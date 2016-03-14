package scalexcel

import scala.reflect.runtime.universe._
import org.apache.poi.ss.usermodel._
import scala.util.{Failure, Success, Try}
import ECell._
import java.util.Date
import org.apache.poi.hssf.util.CellReference

case class Formula(formula: String)

trait CellCoords {
  val rowIdx: Int
  val colIdx: Int
}

trait CellOps extends CellCoords {
  def get[T: TypeTag](implicit sheet: Sheet): T = getOpt.get
  def getOpt[T: TypeTag](implicit sheet: Sheet): Option[T] = getByType(typeOf[T], rowIdx, colIdx).asInstanceOf[Option[T]]
  def getOrElse[T: TypeTag](default: => T)(implicit sheet: Sheet): T = getOpt.getOrElse(default)

  def set(value: Double)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value)
  def set(value: String)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value)
  def set(value: Date)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value)
  def set(value: Boolean)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value)
  def set(value: Formula)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value)

  def set(value: Any)(implicit sheet: Sheet) = ECell(rowIdx, colIdx, value) // only for usage in collection of [Any]
}

class ECell(row: Int, col: Int) extends CellOps with RowOps with TableOps {
  lazy val rowIdx = row
  lazy val colIdx = col
}

object ECell {
  def apply(row: Int, col: Int): ECell = new ECell(row, col)
  def apply(cellRef: String): ECell = {
    val ref = new CellReference(cellRef)
    ECell(ref.getRow, ref.getCol)
  }

  def apply(row: Int, col: Int, value: Double)(implicit sheet: Sheet) = new DoubleCell(row, col).set(value)
  def apply(row: Int, col: Int, value: String)(implicit sheet: Sheet) = new StrCell[String](row, col).set(value)
  def apply(row: Int, col: Int, value: Date)(implicit sheet: Sheet) = new DateCell[Date](row, col).set(value)
  def apply(row: Int, col: Int, value: Boolean)(implicit sheet: Sheet)= new BoolCell[Boolean](row, col).set(value)
  def apply(row: Int, col: Int, value: Formula)(implicit sheet: Sheet)= new FrmlCell[String](row, col).set(value.formula)

  def apply(row: Int, col: Int, value: Any)(implicit sheet: Sheet) = value match {
    case v: Double => new DoubleCell(row, col).set(v)
    case Some(v: Double) => new DoubleCell(row, col).set(v)
    case v: String => new StrCell[String](row, col).set(v)
    case Some(v: String) => new StrCell[String](row, col).set(v)
    case v: Date => new DateCell[Date](row, col).set(v)
    case Some(v: Date) => new DateCell[Date](row, col).set(v)
    case v: Boolean => new BoolCell[Boolean](row, col).set(v)
    case Some(v: Boolean) => new BoolCell[Boolean](row, col).set(v)
    case f: Formula => new FrmlCell[String](row, col).set(f.formula)
    case Some(f: Formula) => new FrmlCell[String](row, col).set(f.formula)
    case None => new EmptyCell[String](row, col).set("") // do nothing
//    case ::(head, tail) =>
//      ???
    case _  => throw new IllegalArgumentException("Type " + value.getClass +
      " is not supported by this library, please use Double, String, Date, Boolean or Formula instead")
  }

  // this method is used when instantiating a cell from a list
  private[scalexcel] def getByType(atype: Type, rowIdx: Int, colIdx: Int)(implicit sheet: Sheet): Option[Any] = {
    if (atype =:= typeOf[Double] || atype =:= typeOf[Option[Double]]) new DoubleCell(rowIdx, colIdx).get
    else if (atype =:= typeOf[String] || atype =:= typeOf[Option[String]]) new StrCell[String](rowIdx, colIdx).get
    else if (atype =:= typeOf[Boolean] || atype =:= typeOf[Option[Boolean]]) new BoolCell[Boolean](rowIdx, colIdx).get
    else if (atype =:= typeOf[Date] || atype =:= typeOf[Option[Date]]) new DateCell[Date](rowIdx, colIdx).get
    else if (atype =:= typeOf[Formula] || atype =:= typeOf[Option[Formula]]) new FrmlCell[String](rowIdx, colIdx).get
    else if (atype =:= typeOf[Nothing]) {
      // No type was given as value[T], getting a cell type in runtime
      val cell = poiCell(rowIdx, colIdx)
      if (cell == null) new EmptyCell[String](rowIdx, colIdx).get
      else cell.getCellType match {
        case Cell.CELL_TYPE_BLANK => new EmptyCell[String](rowIdx, colIdx).get
        case Cell.CELL_TYPE_BOOLEAN => new BoolCell[Boolean](rowIdx, colIdx).get
        case Cell.CELL_TYPE_ERROR => new ErrCell[Byte](rowIdx, colIdx).get
        case Cell.CELL_TYPE_FORMULA => new FrmlCell[String](rowIdx, colIdx).get // TODO:??
        case Cell.CELL_TYPE_NUMERIC =>
          if (isDateCell(rowIdx, colIdx)) new DateCell[Date](rowIdx, colIdx).get
          else new DoubleCell(rowIdx, colIdx).get
        case Cell.CELL_TYPE_STRING => new StrCell[String](rowIdx, colIdx).get
      }
    }
    else throw new IllegalArgumentException("Type " + atype + " is not supported by the library")
  }

  def poiCell(rowIdx: Int, cellIdx: Int)(implicit sheet: Sheet): Cell = {
    val row = getOrCreateRow(sheet, rowIdx)
    getOrCreateCell(row, cellIdx)
  }

  // private

  private[scalexcel] def handleEx[T](res: Try[Any], row: Int, col: Int): Option[T] = res match {
    case Success(v) => if (v == null) None else Some(v).asInstanceOf[Option[T]]
    case Failure(ex) => errMsg(row, col, ex); None
  }

  private[scalexcel] def errMsg(row: Int, col: Int, ex: Throwable) = {
    val err = "Error reading cell at (" + row + ", " + col + "), ex: " + ex + ", returning None"
    println(err); // TODO: refactor this
    err
  }

  private[scalexcel] def getOrCreateRow(sheet: Sheet, rowIdx: Int): Row = {
    val row = sheet.getRow(rowIdx)
    if (row == null) sheet.createRow(rowIdx) else row
  }

  private[scalexcel] def getOrCreateCell(row: Row, cellIdx: Int): Cell = {
    val cell = row.getCell(cellIdx)
    if (cell == null) row.createCell(cellIdx) else cell
  }

  private def getOrCreateCell(rowIdx: Int, cellIdx: Int)(implicit sheet: Sheet): Cell = {
    val row = getOrCreateRow(sheet, rowIdx)
    getOrCreateCell(row, cellIdx)
  }

  private[scalexcel] def createDateStyle(sheet: Sheet) = {
    val wb = sheet.getWorkbook
    val helper = wb.getCreationHelper
    val style = wb.createCellStyle()
    style.setDataFormat(helper.createDataFormat().getFormat("m/d/yy"))
    style
  }

  private[scalexcel] def isDateCell(row: Int, col: Int)(implicit sheet: Sheet) = {
    DateUtil.isCellDateFormatted(getOrCreateCell(row, col))
  }
  
  private[scalexcel] def getStringForced(row: Int, col: Int)(implicit sheet: Sheet) = {
    val cell = poiCell(row, col)
    val cellType = cell.getCellType
    cell.setCellType(Cell.CELL_TYPE_STRING)

    val str = cell.getStringCellValue // TODO: orElse might return a wrong result
    cell.setCellType(cellType)

    str
  }
}

abstract class BaseCell[T: TypeTag](row: Int, col: Int) {
  def get(implicit sheet: Sheet): Option[T]
  def set(value: T)(implicit sheet: Sheet): BaseCell[T]
}

case class DoubleCell(row: Int, col: Int) extends BaseCell[Double](row, col) {
  type T = Double
  override def get(implicit sheet: Sheet): Option[T] = {
    val cell = poiCell(row, col)
    val res = handleEx[T](Try(cell.getNumericCellValue), row, col)
    val finalRes = res match {
      case None => None
      case Some(v: Double) => //Some(v)
        if (v != 0.0) Some(v)
        else { // v == 0.0 - it might be empty, but poi returns 0.0 even for empty cells
          val cellType = cell.getCellType
          cell.setCellType(Cell.CELL_TYPE_STRING)
          val str = Try(cell.getStringCellValue).getOrElse("") // TODO: orElse might return a wrong result
          cell.setCellType(cellType)

          if (str == "") None // empty cell
          else Some(v) // value of 0.0
        }
    }
    val r = finalRes.asInstanceOf[Option[T]]
    r
  }
  override def set(value: T)(implicit sheet: Sheet): DoubleCell = { poiCell(row, col).setCellValue(value.asInstanceOf[T]); this }
}
case class StrCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = handleEx(Try(poiCell(row, col).getStringCellValue), row, col)
  override def set(value: T)(implicit sheet: Sheet): StrCell[T] = { poiCell(row, col).setCellValue(value.asInstanceOf[String]); this }
}
case class BoolCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = handleEx(Try(poiCell(row, col).getBooleanCellValue), row, col)
  override def set(value: T)(implicit sheet: Sheet): BoolCell[T] = { poiCell(row, col).setCellValue(value.asInstanceOf[Boolean]); this }
}
case class DateCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = handleEx(Try(poiCell(row, col).getDateCellValue), row, col)
  override def set(value: T)(implicit sheet: Sheet): DateCell[T] = {
    val cell = poiCell(row, col)
    cell.setCellValue(value.asInstanceOf[Date])
    cell.setCellStyle(createDateStyle(cell.getSheet))
    this
  }
}
case class ErrCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = handleEx(Try(poiCell(row, col).getErrorCellValue), row, col)
  override def set(value: T)(implicit sheet: Sheet): ErrCell[T] = { poiCell(row, col).setCellErrorValue(value.asInstanceOf[Byte]); this }
}
case class EmptyCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = { None }
  override def set(value: T)(implicit sheet: Sheet): EmptyCell[T] = { this } // TODO: Change the type if needed
}
case class FrmlCell[T: TypeTag](row: Int, col: Int) extends BaseCell[T](row, col) {
  override def get(implicit sheet: Sheet): Option[T] = handleEx(Try(poiCell(row, col).getCellFormula), row, col)
  override def set(value: T)(implicit sheet: Sheet):FrmlCell[T] = { poiCell(row, col).setCellFormula(value.asInstanceOf[String]); this } // TODO: errors expected.
}


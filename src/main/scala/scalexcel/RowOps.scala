package scalexcel

import scala.reflect.runtime.universe._
import org.apache.poi.ss.usermodel.Sheet
import scala.util.{Failure, Success, Try}

/**
 * Contains Row-related operations
 */
trait RowOps extends CellCoords {

  def getRowAsList[T](implicit T: TypeTag[T], sheet: Sheet): List[Any] = {
    if (typeOf[T] =:= typeOf[Nothing])
      throw new IllegalArgumentException("a type parameter should be specified in square brackets []")
    else {
      val nameTypeMap = Reflection.caseClassParams[T] // TypeTag[Nothing]
      var colNum = colIdx
      val values = for ((name, atype) <- nameTypeMap) yield {
        val value = ECell.getByType(atype, rowIdx, colNum)
        colNum = colNum + 1
        value
      }
      Reflection.unwrapOpt[T](values.toSeq).toList
    }
  }

  def getRow[T](implicit T: TypeTag[T], sheet: Sheet): T = Reflection.newObject[T](getRowAsList[T])

  def getRowOpt[T](implicit T: TypeTag[T], sheet: Sheet): Option[T] = Try(getRow) match {
    case Success(v) => Some(v)
    case Failure(ex) => None
  }

  def setRow(values: Seq[Any])(implicit sheet: Sheet): Seq[Any] = {
    var colNum = colIdx
    values.foreach(v => {
      ECell(rowIdx, colNum).set(v)
      colNum += 1
    })
    values
  }

  def setRow(values: List[Any])(implicit sheet: Sheet): List[Any] = {
    setRow(values.toSeq)
    values
  }

  def setRow2(values: Any *)(implicit sheet: Sheet) = {
    setRow(values.toSeq)
    values
  }

  /**
   * Gets any case class and set its values to excel in a row, as specified in default contructor
   *
   * @param caseClass case class instance to save to excel
   * @param sheet implicit sheet
   * @tparam T type of the case class (or any Product descendant) specified in []
   * @return a case class (or product descendant) saved to an excel
   */
  def setRow[T <: Product](caseClass: T)(implicit sheet: Sheet): T = {
    setRow(caseClass.productIterator.toSeq)
    caseClass
  }

}

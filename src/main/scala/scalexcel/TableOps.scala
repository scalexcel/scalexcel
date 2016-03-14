package scalexcel

import scala.reflect.runtime.universe._
import scala.collection.mutable
import org.apache.poi.ss.usermodel.Sheet

/**
 * Contains operations related to a table
 */
trait TableOps extends CellCoords {
  type Row = Seq[Any]

//  def setTable(values: Seq[Row])(implicit sheet: Sheet): Seq[Row] = {
//    var rowNum = rowIdx
//    for (list <- values) {
//      ECell(rowNum, colIdx) setRow list
//      rowNum += 1
//    }
//    values
//  }

  def setTable(values: Seq[Product])(implicit sheet: Sheet): Seq[Product] = {
    var rowNum = rowIdx
    for (list <- values) {
      ECell(rowNum, colIdx) setRow list
      rowNum += 1
    }
    values
  }



    /**
   * Reads a table of case classes.
   *
   * @param T     the type of the case class (or any Product derived class)
   * @param sheet implicit sheet
   * @tparam T    the type of the case class (or any Product derived class)
   * @return      a list of instances of case class, provied in the []
   */
  def getTable[T <: Product](implicit T: TypeTag[T], sheet: Sheet): Seq[T] = {
    // TODO refactor it to be scalish
    var rowNum = rowIdx
    var records = mutable.MutableList[T]()
    var done = false

    do
      ECell(rowNum, colIdx).getRowOpt[T] match {
        case Some(record) =>
          records += record
          rowNum += 1
        case None => done = true
    } while (!done)

    records.toSeq
  }
}


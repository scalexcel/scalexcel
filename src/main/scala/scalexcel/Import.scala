package scalexcel

/**
 * Contains implicit functions for excel library
 */
object Import {
  implicit def tuple2toECell(tuple: (Int, Int)): ECell = ECell(tuple._1, tuple._2)
  implicit def stringToECell(cellRef: String): ECell = ECell(cellRef)
}

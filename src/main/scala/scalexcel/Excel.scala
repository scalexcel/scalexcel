package scalexcel

import org.apache.poi.ss.usermodel._
import java.io.{OutputStream, FileOutputStream, FileInputStream, InputStream}
import org.apache.poi.xssf.usermodel.{XSSFFormulaEvaluator, XSSFWorkbook}

object Excel {
  implicit def tuple2toECell(tuple: (Int, Int)): ECell = ECell(tuple._1, tuple._2)
  implicit def stringToECell(cellRef: String): ECell = ECell(cellRef)

  /**
   * Opens an excel file for reading and writing. Please have in mind that only .xlsx files are supported
   * <p>Example of usage:
   * <pre>
   * {@code
    open(xlsxTemplate, fileOut) { implicit workbook =>
      removeSheet(export)

      sheet(export) { implicit sheet =>
        (0, 0) setRow header.toList
        for (i <- 0 until records.length) {
          val rowNum = i + 1
          (rowNum, 0) setRow records(i).productIterator.toList // convert a case class of List of values
        }
      }
    }
   * </pre>
   *
   * @param fileName  the name of the excel file to open.
   * @param block     the code block to be executed inside of the this method. See the example above in the description
   * @return          Unit
   */
  def open[T](fileName: String)(block: Workbook => T): T = {
    excel(fileName, fileName)(block)
  }

  class _Workbook(poiWb: Workbook) {
    def sheet[T](sheetName: String)(block: Sheet => T)(implicit workbook: Workbook): T = {
      val sh = getOrCreateSheet(workbook, sheetName)
      try {
        block(sh)
      } finally {}
    }

    def removeSheet(sheetName: String)(implicit wb: Workbook) = wb.removeSheetAt(wb.getSheetIndex(sheetName))
    def removeSheet(index: Int)(implicit wb: Workbook) = wb.removeSheetAt(index)
  }

  def workbook[T](fileName: String): _Workbook = {
    implicit val wb = WorkbookFactory.create(new FileInputStream(fileName))
    new _Workbook(wb)
  }

  def excel[T](inp: InputStream, outFileName: String)(block: Workbook => T): T = {
    val wb = WorkbookFactory.create(inp)
    workbookBlock(wb, outFileName)(block)
  }

  def excel[T](inp: InputStream)(block: Workbook => T): T = {
    val wb = WorkbookFactory.create(inp)
    workbookBlock(wb)(block)
  }

  def excel[T](fileName: String, outFileName: String)(block: Workbook => T): T = {
    excel(new FileInputStream(fileName), outFileName)(block)
  }

  def newExcel[T](fileName: String)(block: Workbook => T): T = {
    val wb = new XSSFWorkbook()
    workbookBlock(wb, fileName)(block)
  }

  def sheet[T](sheetName: String)(block: Sheet => T)(implicit workbook: Workbook): T = {
    val sh = getOrCreateSheet(workbook, sheetName)
    try {
      block(sh)
    } finally {
    }
  }

  def sheet[T](fileName: String, sheetName: String)(block: Sheet => T): T = {
    open(fileName) { implicit wb =>
      sheet(sheetName)(block)
    }
  }

  def removeSheet(sheetName: String)(implicit wb: Workbook) = wb.removeSheetAt(wb.getSheetIndex(sheetName))
  def removeSheet(index: Int)(implicit wb: Workbook) = wb.removeSheetAt(index)

  // private methods

  private def getOrCreateSheet(workbook: Workbook, sheetName: String) = {
    val sh = workbook.getSheet(sheetName)
    if (sh == null) workbook.createSheet(sheetName) else sh
  }

  private def workbookBlock[T](wb: Workbook, outFileName: String)(block: Workbook => T): T = {
    workbookBlock(wb, new FileOutputStream(outFileName))(block)
  }

  private def workbookBlock[T](wb: Workbook, os: OutputStream)(block: Workbook => T): T = {
    try {
      block(wb)
    } finally {
      evaluateAllFormulas(wb)
      wb.write(os)
      os.close()
    }
  }

  private def evaluateAllFormulas(wb: Workbook) = {
    val eval = wb.getCreationHelper.createFormulaEvaluator()
    eval.evaluateAll()
  }

  private def workbookBlock[T](wb: Workbook)(block: Workbook => T): T = {
    try {
      block(wb)
    } finally {
    }
  }
}

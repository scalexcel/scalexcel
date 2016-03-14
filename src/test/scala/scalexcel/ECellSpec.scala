package scalexcel

import org.scalatest.{Matchers, FlatSpec}
import org.apache.poi.xssf.usermodel._
import java.util.{GregorianCalendar, Date}
import java.io.{FileInputStream, File}

import scalexcel._
import scalexcel.Import._

class ECellSpec extends FlatSpec with Matchers {
  val fileName = "example/TestWorkbook.xlsx"
  val wb = new XSSFWorkbook(new FileInputStream(fileName))
  val sheetName = "MySheet1"

  "ECell" should "be created" in {
    implicit val sheet = wb.getSheet(sheetName)

    val c1 = ECell(1, 1)
    c1.getOpt[String] should be (Some("Lorem Ipsum"))
    c1.getOpt should be (Some("Lorem Ipsum"))

    ECell(1, 0).getOpt[Double] should be (Some(1.0))
    (1, 0).get[Double] should be (1.0)

    (1, 0).getOpt[Double] should be (Some(1.0))
    val v = (1, 0).getOpt
    println(v)
    v should be (Some(1.0)) // !!! getting a cell type in runtime
    (1, 1).getOpt should be (Some("Lorem Ipsum"))

    // should fail and return None
    (1, 1).getOpt[Double] should be (None) // no double value, only String value
    (1, 1).getOpt[String] should be (Some("Lorem Ipsum"))
    (1, 1).getOpt should be (Some("Lorem Ipsum")) // by default a StringCell is used

    (1, 15).getOpt should be (Some("IF(K2>J2,1,0)"))
    (1, 15).getOpt[Formula] should be (Some("IF(K2>J2,1,0)"))
    (1, 15).getOpt[String] should be (None) // cannot read String from Double cell // TODO: fix it! return String if we ask for it
    (1, 15).getOpt[Double] should be (Some(0.0))

    val d = (1, 5).getOpt[Date] // hired date
    d.get.toString should be ("Mon Nov 19 00:00:00 MSK 2012")
    d.get should be (new GregorianCalendar(2012, 10, 19).getTime)
  }

  it should "be set with a value" in {
    implicit val sheet = wb.getSheet(sheetName)

    val c = ECell(58, 1).set("hello")

    (59, 1) set "hello"
    (60, 0) set 3.14

    ECell(58, 1).getOpt should be (Some("hello"))

    val pi: Double = (60, 0).get
    pi should be (3.14)
  }

  it should "be able to read any value as a String, temporary changing the cell type" in {
    implicit val sheet = wb.getSheet(sheetName)

    val a1: Double = (1, 0).get
    a1 should be (1.0)
    
    val a1str = ECell.getStringForced(1, 0)
    a1str should be ("1")

    val o1 = (1, 8).getOpt[Double]
    o1 should be (None)
//    o1 should be (Some(0.0))

    val money = ECell.getStringForced(1, 8)
    money should be ("")
  }


}

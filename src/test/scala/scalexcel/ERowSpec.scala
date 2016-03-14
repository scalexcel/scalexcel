package org.d5i.excel

import org.scalatest._
import java.util.Date
import org.apache.poi.xssf.usermodel._
import java.io.FileInputStream

import scalexcel._
import scalexcel.Excel._

case class PeopleSal2(num: Double, name: String, role: String, location: String, team: String, startDate: Date, probationEnd: Date, endDate: Option[Date],
                      probationSal: Option[Double], currentSal: Double, salary: Option[Double], comment: String, wasInGland: String, perfApr: String, fromIM: String, column1: Double)

case class PeopleInfo(name: String, role: String, location: String, team: String, startDate: Option[Date],
                      probEndDate: Option[Date], endDate: Option[Date])

case class PeopleInfo3(name: Option[String], role: Option[String], location: String, team: String, startDate: Option[Date],
                       probEndDate: Option[Date], endDate: Option[Date])

class ERowSpec extends FlatSpec with Matchers {
  val fileName = "example/TestWorkbook.xlsx"
  val sheetName = "MySheet1"

  "ERow" should "have an ability to read a full row as a list" in {

    val wb = new XSSFWorkbook(new FileInputStream(fileName))
    implicit val sheet = wb.getSheet(sheetName)

    val res = (1, 1).getRowAsList[PeopleInfo]
    println(res)

    res(0) should be("Lorem Ipsum")

    val res2 = ECell(1, 0).getRowAsList[PeopleSal2]
    println(res2)

    res2(1) should be("Lorem Ipsum")
  }

  it should "have an ability to read a full row as an object" in {
    val inputStream = new FileInputStream(fileName)
    //val inputStream = getClass.getResourceAsStream("/" + fileName)

    excel(inputStream) { implicit workbook =>
      sheet(sheetName) { implicit sheet =>
        val res = (1, 1).getRow[PeopleInfo]
        println(res)
        res.name should be("Lorem Ipsum")

        val res1 = ECell(1, 1).getRow[PeopleInfo3]
        println(res1)
        res1.name should be(Some("Lorem Ipsum"))

        val vp = (1, 0).getRow[PeopleSal2]
        println(vp)
        vp.name should be("Lorem Ipsum")
        vp.probationSal should be(None)

        "A4".getRow[PeopleSal2].probationSal should be(Some(2000.0))

        val res3 = (3, 0).getRow[PeopleSal2]
        println(res3)
        res3.name should be("Lorem Ipsum3")
      }
    }

  }

  it should "be able to write a row of List" in {

    newExcel("out/TestSheetResult.xlsx") { implicit workbook =>
      sheet("Test") { implicit sheet =>
        val list = List(1.0, "222", new Date())
        val l2 = "test sdfs" :: list
        (4, 0) setRow list
        (5, 0) setRow l2
        (6, 0) setRow list ++ l2

        ECell(7, 0).setRow(list ++ l2)
      }
    }
  }

  it should "have new way of opening workbook" in {

    newExcel("out/SUPER OUT.xlsx") { implicit wb =>
      sheet("Test") { implicit sheet =>
        val list = List(1.0, "222", new Date())
        val l2 = "test sdfs" :: list

        (4, 0) setRow list
        (5, 0) setRow l2
        (6, 0) setRow list ++ l2

        ECell(7, 0).setRow(list ++ l2)
      }
    }

    excel(new FileInputStream("out/SUPER OUT.xlsx"), "out/SUPER OUT - RES.xlsx") { wb =>
      implicit val sheet = wb.getSheet("Test")

      (9, 0) setRow List("UPDATED FILE")
    }

    excel("out/SUPER OUT.xlsx", "out/SUPER OUT - 3.xlsx") { implicit wb =>
      implicit val sheet = wb.getSheet("Test")

      (9, 0) setRow List("UPDATED FILE")
    }

    excel("out/SUPER OUT.xlsx", "out/SUPER OUT.xlsx") { wb =>
      implicit val sheet = wb.getSheet("Test")

      (10, 0) setRow List("THE SAME UPDATED FILE")
    }

    open("out/SUPER OUT.xlsx") { wb =>
      implicit val sheet = wb.getSheet("Test")

      (10, 0) setRow List("THE SAME UPDATED FILE", "EVEN two timess")
    }
  }

  it should "have a DSL for working with sheet" in {
    newExcel("out/Sheet DSL.xlsx") { implicit wb =>
      sheet("Test") { implicit sheet =>
        (0, 0) setRow List("first col", 111.0, true)
      }
    }

    sheet("out/Sheet DSL.xlsx", "Test") { implicit sheet =>
      (1, 0) setRow List("first col", 222.0, true)
    }

    sheet("out/Sheet DSL.xlsx", "Test") { implicit sheet =>
      (2, 0) setRow List("third row", 333.0, true)
    }

    // TODO:
    // Decrease amoount of interfaces methods
    //
    //    create workbook "Employees.xlsx" sheet "Test" {
    //
    //    }
    //    open workbook "Employees.xlsx" sheet "Test" {
    //
    //    }

    // TODO: use "active data structure" concept
    // https://resources.mpi-inf.mpg.de/conferences/adfocs-03/Slides/Meyer_1.pdf

    // val res = below("First Name") getRow[Employee]
    //


  }

}

package scalexcel

import java.util.Date

import org.scalatest._

import scalexcel._
import scalexcel.Excel._

case class Employee(name: String, salary: Double, hasPet: Boolean, hired: Date, fired: Option[Date])

class ExcelSpec extends FlatSpec with Matchers {
  "Excel" should "have an example: 'how to open an excel file for reading and read one value'" in {

    val date = new Date()

    newExcel("out/simple.xlsx") { implicit workbook =>
      sheet("testSheet") { implicit sheet =>
        (0, 0) set "test string"
        ECell(0, 0) set "test string"
        (1, 0) set 273.45
        (2, 0) set date
        (3, 0) set true
        (4, 0) set Formula("1+1")

        "A10" set "another test string"
//        A10 set "another test string"

        ECell("A10") set "another test string"
      }
    } // the file is saved to disk

    val res = sheet("out/simple.xlsx", "testSheet") { implicit sheet =>
      val str: String = (0, 0).get    // to be type safe, scala needs to know the type of the result
      val st1 = (0, 0).get[String]    // or this way

      val dbl: Double = (1, 0).get    // another example
      val db2 = (1, 0).get[Double]
      val dbOpt = (1, 0).getOpt[Double]
      val db3 = "A2".getOrElse(0.0)
      db3 should be (273.45)

      val strOpt = (0, 0).getOpt    // in compile time strOpt is of type Option[Nothing], in runtime, the object is of type Option[String]
      // val s1 = strOpt.get          // would be an exception: java.lang.String cannot be cast to scala.runtime.Nothing$
      val ss1: String = strOpt.get  // so do this way
      val sss  = strOpt.getOrElse("") // or this way, specifying a default value from with a type can be obtained

      val dblOpt = (1, 0).getOpt
      val date = (2, 0).getOpt
      val bool = (3, 0).getOpt
      val formula = (4, 0).getOpt
      val frmlVal = (4, 0).getOpt[Double] // if we know that the value of formula is of type Double (NUMERIC)

      val another = "A10".getOpt
      another should be (Some("another test string"))

      val empty = (100, 0).getOpt // None is returned

      // more detailed case with types
      val str1 = (0, 0).getOpt[String] // if you know the type you would like to get
      val str3: Option[String] = (0, 0).getOpt // this is the real return type

      val dbl1 = (1, 0).getOpt[Double]
      val dbl3: Option[Double] = (1, 0).getOpt

      val date1 = (2, 0).getOpt[Date]
      val date3: Option[Date] = (2, 0).getOpt

      val bool1 = (3, 0).getOpt[Boolean]
      val bool3: Option[Boolean] = (3, 0).getOpt

      val formula1 = (4, 0).getOpt[Formula]
      val formula3: Option[Formula] = (4, 0).getOpt // it will get a formula, not its value

      // example of returning from 'sheet' block

      List(str, dbl, date, bool, formula, frmlVal)
    }

    //res should be (List("test string", 273.45, date, true, "1+1", 2.0))
  }

  it should "have an example: 'reading and writing case classes - most powerful way of working with library'" in {
//    import org.d5i.excel._

    // first, let us create a data
    newExcel("out/simple2.xlsx") { implicit workbook =>
      sheet("Employees") { implicit sheet =>
        ECell(0, 0).setRow(List("Peter", 120000.0, true, new Date(), None, "and", "lots of ", "other fields", Some("option")))

        ECell(0, 0) setRow List("Peter", 120000.0, true, new Date(), None, "and", "lots of ", "other fields", Some("option"))
        (1, 0) setRow List("Mike", 110001.0, false, new Date(), new Date)
        (2, 0) setRow Employee("Vanessa", 110002.0, hasPet = true, new Date(), None).productIterator.toList
        (3, 0) setRow Employee("Alex", 110003.0, hasPet = true, new Date(), None) // only top-level case classes are  for now
        (4, 0) setRow ("Alex", 110004.0, true, new Date(), None) // one parameter - Product

      }
    } // the file is saved to disk and formulas are evaluated

    val res = sheet("out/simple2.xlsx", "Employees") { implicit sheet => // we can read a sheet in an excel file!
      val peter = (0, 0).getRow[Employee]
      val mike = (1, 0).getRow[Employee]
      val vanessaList = (2, 0).getRowAsList[Employee]

      // the same with types
      val peter1: Employee = (0, 0).getRow[Employee]
      val peter100Opt = (0, 0).getRowOpt[Employee] // None is returned, type is Option[Employee]
      val mike1: Employee = (1, 0).getRow[Employee]
      val vanessaList1: List[Any] = (2, 0).getRowAsList[Employee]

      // we can return a value from block
      (peter, mike, vanessaList)
    }

    println(res)
  }

  it should "have an example: 'reading and writing a full table'" in {

    newExcel("out/simple3.xlsx") { implicit workbook =>
      sheet("Tables") { implicit sheet =>
//        val records = List(
//          List("Peter1", 100000.0, true, new Date, None),
//          List("Peter2", 110000.0, true, new Date, Some(new Date)),
//          List("Peter3", 120000.0, true, new Date, None),
//          List("Peter4", 130000.0, true, new Date, Some(new Date)))

        val records = List(
          Employee("Peter1", 100000.0, true, new Date, None),
          Employee("Peter2", 110000.0, true, new Date, Some(new Date)),
          Employee("Peter3", 120000.0, true, new Date, None),
          Employee("Peter4", 130000.0, true, new Date, Some(new Date)))

        "A10" setTable records
        (0, 0) setTable records // at "A1"
        "AA10" setTable records
      }
    }

    sheet(fileName = "out/simple3.xlsx", sheetName = "Tables") { implicit sheet =>
      val empty = (4, 0).getRowOpt[Employee] // returns an Option[Employee] = None, because the row is empty
      empty should be (None)

      val records1 = (0, 0).getTable[Employee] // the type is necessary, because it has a types of all the table columns
      records1(0).name should be ("Peter1")
      records1(0).fired should be (None)

      records1.foreach(println)


      val records2 = "A1".getTable[Employee]

      println(records2)
    }


  }
}







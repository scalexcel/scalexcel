
name := "scalexcel"

version := "0.1.0"

scalaVersion := "2.11.7"

libraryDependencies ++= Seq(
    "org.apache.poi" % "poi" % "3.9",
    "org.apache.poi" % "poi-ooxml" % "3.9")

libraryDependencies += "org.scala-lang" % "scala-reflect" % scalaVersion.value

//libraryDependencies += "org.scalactic" %% "scalactic" % "2.2.6"
libraryDependencies += "org.scalatest" %% "scalatest" % "2.2.6" % "test"


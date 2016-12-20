name := "scalexcel"

organization := "com.github.scalexcel"

version := "0.1.6"

scalaVersion := "2.11.7"

libraryDependencies ++= Seq(
    "org.scala-lang" % "scala-reflect" % scalaVersion.value,
    "org.apache.poi" % "poi" % "3.9",
    "org.apache.poi" % "poi-ooxml" % "3.9",
    "org.scalatest" %% "scalatest" % "2.2.6" % "test")


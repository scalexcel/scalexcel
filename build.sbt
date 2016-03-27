name := "scalexcel"

organization := "scalexcel"

version := "0.1.0"

scalaVersion := "2.11.7"

libraryDependencies ++= Seq(
    "org.scala-lang" % "scala-reflect" % scalaVersion.value,
    "org.apache.poi" % "poi" % "3.9",
    "org.apache.poi" % "poi-ooxml" % "3.9",
    "org.scalatest" %% "scalatest" % "2.2.6" % "test")

// publish related
sonatypeProfileName := "com.github.scalexcel"

publishMavenStyle := true

publishArtifact in Test := false

publishTo := {
  val nexus = "https://oss.sonatype.org/"
  if (isSnapshot.value)
    Some("snapshots" at nexus + "content/repositories/snapshots")
  else
    Some("releases"  at nexus + "service/local/staging/deploy/maven2")
}

pomIncludeRepository := { _ => false }

pomExtra := (
  <url>https://github.com/scalexcel/scalexcel</url>
    <licenses>
      <license>
        <name>BSD-style</name>
        <url>http://www.opensource.org/licenses/bsd-license.php</url>
        <distribution>repo</distribution>
      </license>
    </licenses>
    <scm>
      <url>https://github.com/scalexcel/scalexcel.git</url>
      <connection>https://github.com/scalexcel/scalexcel.git</connection>
    </scm>
    <developers>
      <developer>
        <id>scalexcel</id>
        <name>scalexcel</name>
        <url>https://github.com/scalexcel</url>
      </developer>
    </developers>)
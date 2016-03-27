// publish related
//sonatypeProfileName := "com.github.scalexcel"
// it is not needed, because the organization in the build.sbt should be com.github.scalexcel, otherwise an exception occures

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
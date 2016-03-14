package scalexcel

import scala.reflect.runtime.universe._
import scala.collection.immutable.ListMap

import scala.reflect.runtime.{universe=>ru}

// http://www.michaelpollmeier.com/fun-with-scalas-new-reflection-api-2-10/
object Reflection {

  def unwrapOpt[T: TypeTag](args: Seq[Any]): Seq[Any] = {
    def unwrap(arg: Any, atype: Type) = if (atype <:< typeOf[Option[Any]]) arg else arg match {
      case Some(v) => v
      case None => None
    }
    val types = caseClassParamTypes[T]
    val res = (args zip types) map { case (arg, atype) => unwrap(arg, atype) }
    res
  }

  /**
   * Instantiates a new object based on the given list of constructor arguments and its type.
   *
   * @param args a list of parameters for instantiating an object
   * @tparam T type of the object to be created
   * @return a new instance of an object
   */
  def newObject[T: ru.TypeTag](args: Seq[Any]): T = {
    val typeTag = ru.typeTag[T]
    val runtimeMirror = ru.runtimeMirror(getClass.getClassLoader)

    val classSymbol = typeTag.tpe.typeSymbol.asClass
    val classMirror = runtimeMirror.reflectClass(classSymbol)

    val constructorSymbol = typeTag.tpe.declaration(ru.nme.CONSTRUCTOR).asMethod
    val constructorMirrror = classMirror.reflectConstructor(constructorSymbol)

    // to prevent exception about argument type mismatch, need to add checking here

    // TODO
//    try {
      constructorMirrror(args: _*).asInstanceOf[T]
//    } catch {
//      case ex: IllegalArgumentException => System.out.println(ex.getMessage +
//        "\nOne of the parameters types in your case class does not match with real value type in Excel. " +
//        "\nYou probably forgot to set Option[] for nullable parameters. Please review it and correct the type. \n"
//        + "(expected, actual) types: " + caseClassParamTypes[T].zip(listParamTypes(args)) + "\n" + "args: " + args)
//        ex.printStackTrace(System.out)
//        throw ex
//      case ex: Exception => System.out.println(ex.getMessage, ex)
//        ex.printStackTrace(System.out)
//        throw ex
//    }

    // TODO: add library logger
  }

  def listParamTypes(args: Seq[Any]) = {
    for (arg <- args) yield arg.getClass
  }

  /**
   * Returns a map from formal parameter names to types, containing one
   * mapping for each constructor argument.  The resulting map (a ListMap)
   * preserves the order of the primary constructor's parameter list.
   *
   * case class Foo(foo: String, bar: Int)
   * scala.collection.immutable.ListMap[String,reflect.runtime.universe.Type] = Map(foo -> String, bar -> scala.Int)
   */
  def caseClassParams[T: TypeTag]: ListMap[String, Type] = {
    val tpe = typeOf[T]
    val constructorSymbol = tpe.declaration(nme.CONSTRUCTOR)
    val defaultConstructor =
      if (constructorSymbol.isMethod) constructorSymbol.asMethod
      else {
        val ctors = constructorSymbol.asTerm.alternatives
        ctors.map { _.asMethod }.find { _.isPrimaryConstructor }.get
      }

    ListMap[String, Type]() ++ defaultConstructor.paramss.reduceLeft(_ ++ _).map {
      sym => sym.name.toString -> tpe.member(sym.name).asMethod.returnType
    }
  }

  /**
   * Returns types of parameters in a given case class.
   *
   * @tparam T type of the case class to analyze
   * @return a sequence of types as specified in constructor.
   */
  def caseClassParamTypes[T: TypeTag]: Seq[Type] = {
    val nameTypes = caseClassParams[T]
    val types = for ((name, atype) <- nameTypes) yield atype
    types.toSeq
  }
}

case class Thing(a: Int, b: String, c: Option[Double])

object Test3 extends App {
  val v = Vector(1, "str", Some(234.3))
  val thing: Thing = Reflection.newObject[Thing](v)
  println(thing)
  // prints: Thing(1,str,7.3)

  val v2 = Vector(5, "456", None)
  //)
  val thing2: Thing = Reflection.newObject[Thing](v2)
  println(thing2)
  // prints: Thing(1,str,7.3)

  val res = Reflection.unwrapOpt[Thing](List(Some(1), Some("test"), None))
  println(res)
  val t3 = Reflection.newObject[Thing](res)
  println(t3)

  val t = Thing(1, "b", Some(23.4))
  val list = t.productIterator.toList
  println(list)
}
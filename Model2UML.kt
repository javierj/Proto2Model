import org.apache.poi.xslf.usermodel.*
import java.io.FileInputStream
import java.io.File

fun demo1() {
    val s1 = Screen("Main")
    val s2 = Screen("Search")
    val s3 = Screen("NewElement")

    val l1 = Link(s1, s2)
    val l2 = Link(s1, s3)

    var b1: Boolean
    b1 = Link(s1, s2) == l1
    println("$b1")
    b1 = Link(s1, s2).equals(l1)
    println("$b1")
    b1 = Link(s1, s2) === l1
    println("$b1")

    val screens = listOf(s1, s2, s3)
    val links = listOf(l1, l2)

    println("@startuml")
    for (s in screens)
        println(plainUML(s))
    for (l in links)
        println(plainUML(l))
    println("@enduml")

}

fun main() {
    println("Hello, world!!!")

    //demo1()


    //val fichero = "DosPanyallasUnBoton.pptx"
    val fichero = "EjemploFakeApp.pptx"
    println("Carcar presentación: " + fichero)
    val ppt = load(fichero)

    val screens = extractScreens(ppt)


    val links = mutableSetOf<Link>()
    for (slide in ppt.slides) {
        val key = slide.slideName.toLowerCase()
        //println("$key: ${slide.slideName} / $screens")

        val orgScreen = screens[key]!!

        for (sh in slide.shapes) {
            val targetScreen = getHyperlink(sh, screens)
            if (isValidLink(orgScreen, targetScreen)) {
                var safeScreen: Screen = Screen("name")
                if (targetScreen != null)
                    safeScreen = targetScreen
                val newLink = addLink(orgScreen, safeScreen, sh)
               /* println("1- Link: $newLink / $links / ${links.contains(newLink)}")
                for(l in links) {
                    println("$l : ${l.equals(newLink)} ")
                }*/
                links.add(newLink)
                //println("2- Link: $newLink / $links / ${links.contains(newLink)}")
            }
        }
    }

    toUMLModel(screens, links)
    println(json(screens, links))
    saveJsonFile(json(screens, links))


}



fun load(file: String): XMLSlideShow {
    return XMLSlideShow(FileInputStream(file))
}

fun extractScreens(ppt: XMLSlideShow): Map<String, Screen> {
    val screens = mutableMapOf<String, Screen>()

    for (slide in ppt.slides) {
        val screen = slide2screen(slide)
        screens[screen.key]= screen
    }

    return screens
}

fun getFirstImageIn(slide: XSLFSlide): String {
    for (sh in slide.shapes) {
        if (sh is XSLFPictureShape) {
            return sh.pictureData.fileName
        }
    }
    return "No image found in slide: ${slide.slideName}"
}

fun slide2screen(slide: XSLFSlide): Screen {
    val s = Screen(slide.slideName)
    s.backgroundImageFilename = getFirstImageIn(slide)
    return s
}

fun getHyperlink(shape: XSLFShape, screens : Map<String, Screen> ): Screen? {
    if (shape is XSLFSimpleShape) {
        // org.apache.poi.xslf.usermodel.XSLFHyperlink
        val hl = shape.hyperlink
        //print(" Hyperlink $hl")
        if (hl != null) {

            val pieces = hl.address.split("/").last()
            val key = pieces.split(".").first().toLowerCase()
            //println(" address ${hl.address}, label: ${hl.label} , screen: $screens[key], key: $key")
            return screens[key]
        }
    }
    return null
}

fun isValidLink(src: Screen?, tgt: Screen?): Boolean {
    if ((src == null) || (tgt == null))
        return false
    if (src.key == tgt.key)
        return false
    return true
}

fun  addLink(src: Screen, tgt: Screen, shape: XSLFShape): Link {
    val link = Link(src, tgt)
    src.addLink(link)
    tgt.addLink(link)
    if (shape is XSLFTextShape) {
        link.desc = shape.text
    }

    return link
}

//---- Model --------------------------

class Screen(val name: String) {
    val key = name.toLowerCase()
    val linkTo = mutableListOf<Link>()
    var backgroundImageFilename = "No"

    fun addLink(link: Link) {
        this.linkTo.add(link)
    }

}

data class Link(val source:Screen, val target: Screen) {
    val key = source.key + "_" + target.key
    var desc = ""
    // necesito un equals - esto no me funciona sin data y con data ya no hace falta
    /*override operator fun equals(other: Any?): Boolean {
        if (other is Link) {
            val b = (source == other.source) && (target == other.target)
            //println("Return $b")
            return b
        }
        println("Return false")
        return false
    }*/
}


//--- Model to UML ----------------------

fun plainUML(screen: Screen): String {
    return "class "+screen.name + " << screen >>"
}

fun plainUML(link: Link): String {
    return link.source.name + " -> " + link.target.name
}

fun toUMLModel(screens: Map<String, Screen>, links: Set<Link>) {
    println("@startuml")
    for (s in screens.values)
        println(plainUML(s))
    for (l in links)
        println(plainUML(l))
    println("@enduml")
}

//--- Model too Json for: -------------------------------

fun json(screen:Screen, x:Int, y:Int, w:Int = 100, h:Int = 80) : String {
    val template = """
        {
		"id": "${screen.key}",
		"type": "ifml.ViewContainer",
		"attributes": {
			"name": "${screen.name}",
			"default": true,
			"landmark": false,
			"xor": false
		},
		"metadata": {
			"graphics": {
				"position": {
					"x": $x,
					"y": $y
				},
				"size": {
					"width": $w,
					"height": $h
				}
			}
		}
	}
    """.trimIndent()



    return template
}

fun json(link:Link, x:Int, y:Int) : String {
   /* val template =  """
       {
       "id": "${link.source.key +"_"+ link.target.key}",
       "type": "ifml.Event",
       "attributes": {
           "name": "${link.name}"
       },
       "metadata": {
           "graphics": {

               "name": {
                   "horizontal": "right-outer",
                   "vertical": "top"
               }
           }
       }
   }
   """.trimIndent()
*/
    /*
    "position": {
                   "x": 210,
                   "y": 150
               },
     */

    var template = """
        {
		"id": "${link.key}",
		"type": "ifml.NavigationFlow",
		"attributes": {
			"bindings": []
		},
		"metadata": {

		}
	}
    """.trimIndent()
    return template
}



fun json_relations(links: Set<Link>):String {
    var template = """ """
    for (l in links) {


        template += """
		{
		"type": "source",
		"flow": "${l.key}",
		"source": "${l.source.key}"
	},
	{
		"type": "target",
		"flow": "${l.key}",
		"target": "${l.target.key}"
	}
    """.trimIndent()
        template += ", \n"
    }
    return template.substring(0, template.length-3)
}

fun json(screens: Map<String, Screen>, links: Set<Link>):String {
    var screens_template = ""
    var x = 10
    var y = 10
    var linex = 0
    for (s in screens.values) {
        screens_template += json(s, x, y) + ", \n"
        x+=300
        linex += 1
        if (linex  > 3) {
            linex = 0
            x = 10
            y += 200
        }

    }

    for (l in links) {
        screens_template += json(l, x, y) + ", \n"
    }

    screens_template = screens_template.substring(0, screens_template.length-3)

    val template = """
        {
	"elements": [$screens_template],
	"relations": [${json_relations(links)} ]
    }
    """.trimIndent()


    return template
}



fun saveJsonFile(jsonText: String) {
    val myfile = File("demo1.json")

    myfile.writeText(jsonText)
}
package com.cave.excelconvertor

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.BufferedReader
import java.io.File
import java.io.FileInputStream

val ANDROID_STUDIO_MAIN_VALUES_PATH =
    ""

fun main() {
    val assetsPath = System.getProperty("user.dir")?.plus("/app/src/main/assets") ?: ""
    start(path = assetsPath)
}

private fun start(path: String) {
    File(path).walkBottomUp().forEach { file ->
        if (file.path.contains(".xls")) {
            val splits = file.path.split("/")
            var fileName = ""
            for ((index, item) in splits.withIndex()) {
                if (index == splits.size - 1) {
                    fileName = item.replace(".xls", "")
                }
            }
            readEXCELFile(path = file.path, fileName = fileName)
        }
    }
}

private fun readEXCELFile(path: String, fileName: String) {
    val src = File(path)
    val fis = FileInputStream(src)
    val wb = HSSFWorkbook(fis)
    val sheet = wb.getSheetAt(0)
    var text = ""
    var errorMessages: String? = null
    val keys = checkForRules(fileName = fileName)
    sheet.forEach {
        val key = sheet.getRow(it.rowNum).getCell(0).richStringCellValue
        if (keys.containsValue(key.toString())) {
            val value = sheet.getRow(it.rowNum).getCell(2).richStringCellValue
            val xmlString = "<string name=\"$key\">$value</string>"
            text += xmlString + "\n"
        } else {
            errorMessages?.let { errorMessages += "\"$key\" not find in main source\n" } ?: run { errorMessages = "" }
        }
    }
    errorMessages?.let {
        writeErrorFile(text = it, path = path)
    }
    writeInFile(text = text, path = path)
}
private fun writeErrorFile(text: String, path: String) {
    val writeFileRouteWithName = path.replace(".xls", "_error.txt")
    File(writeFileRouteWithName).writeText(text)
}

private fun writeInFile(text: String, path: String) {
    val xmlFormat = "<resources>\n$text\n</resources>"
    val writeFileRouteWithName = path.replace(".xls", ".xml")
    File(writeFileRouteWithName).writeText(xmlFormat)
}

private fun checkForRules(fileName: String): Map<Int, String> {
    val filePath =
        "$ANDROID_STUDIO_MAIN_VALUES_PATH/$fileName.xml"
    val keysHashMap = HashMap<Int, String>()
    File(ANDROID_STUDIO_MAIN_VALUES_PATH).walkBottomUp()
        .forEach { file ->
            if (file.path.equals(filePath)) {
                val bufferedReader: BufferedReader = file.bufferedReader()
                val lines = bufferedReader.readLines()
                for ((index, item) in lines.withIndex()) {
                    if (item.contains("<string")) {
                        val key = item.substring(item.indexOf("\"") + 1, item.lastIndexOf("\""))
                        keysHashMap[index] = key
                    }
                }
            }
        }
    return keysHashMap
}

package se.johannalynn.nw

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*

class Main {

    data class Result(val nbr: Int, val points: Int, val time: Int, val errors: Int, val sse: Boolean)

    val participants = mutableMapOf<Int, Pair<String,String>>()
    val searchOne = mutableListOf<Result>()
    val searchTwo = mutableListOf<Result>()
    val searchThree = mutableListOf<Result>()
    val searchFour = mutableListOf<Result>()

    val searchNames = mutableListOf<String>()
    val COLUMNS = arrayOf("#","Förare","Hund","Poäng","Fel","Tid","SSE")

    private fun generateResult(nbr: Int, pointsIn: String, errorsIn: String, timeIn: String, sseIn: String): Result {
        val points = if (pointsIn.isNotEmpty()) pointsIn.toInt() else 0
        val time = if (timeIn.isNotEmpty()) {
            val timePart1 = timeIn.split(":")
            val minutes = timePart1[0].toInt()
            val timePart2 = timePart1[1].split(".")
            val seconds = timePart2[0].toInt()
            val milliseconds = timePart2[1].toInt()

            // println("$minutes : $seconds , $milliseconds")
            minutes * 60 * 1000 + seconds * 1000 + milliseconds
        } else {
            0
        }
        val errors = if (errorsIn.isNotEmpty()) errorsIn.toInt() else 0
        val sse = sseIn.isNotEmpty()

        return Result(nbr, points, time, errors, sse)
    }

    fun loadData(lines: List<String>): Boolean {
        lines.forEachIndexed { index, line ->
            if (index == 0) {
                val cells = line.split(",")
                searchNames.add(cells[3])
                searchNames.add(cells[7])
                searchNames.add(cells[11])
                searchNames.add(cells[15])
            } else if (index > 1) {
                val cells = line.split(",")
                // println(cells)
                val nbr = cells[0].toInt()
                participants[nbr] = Pair(cells[1], cells[2])
                val result1 = generateResult(nbr, cells[3], cells[4], cells[5], cells[6])
                searchOne.add(result1)
                val result2 = generateResult(nbr, cells[7], cells[8], cells[9], cells[10])
                searchTwo.add(result2)
                val result3 = generateResult(nbr, cells[11], cells[12], cells[13], cells[14])
                searchThree.add(result3)
                val result4 = generateResult(nbr, cells[15], cells[16], cells[17], cells[18])
                searchFour.add(result4)
            }
        }
        return true
    }

    fun calcResult() {
        val workbook = XSSFWorkbook()

        printResult(workbook, searchOne, searchNames[0])
        printResult(workbook, searchTwo, searchNames[1])
        printResult(workbook, searchThree, searchNames[2])
        printResult(workbook, searchFour, searchNames[3])

        val fileName = "result.xlsx"
        println("Write result to: ${fileName}")
        val fileOut = FileOutputStream(fileName)
        workbook.write(fileOut)
        fileOut.close()
    }

    fun printResult(workbook: XSSFWorkbook, searchList: List<Result>, name: String) {
        val comparator = compareByDescending<Result> { it.points }.thenBy { it.errors }.thenBy { it.time }
        val searchResult = searchList.sortedWith(comparator)

        val createHelper = workbook.getCreationHelper()
        val sheet = workbook.createSheet(name)

        val headerFont = workbook.createFont()
        headerFont.setBold(true)

        val headerCellStyle = workbook.createCellStyle()
        headerCellStyle.setFont(headerFont)

        val headerRow = sheet.createRow(0)
        for (col in COLUMNS.indices) {
            val cell = headerRow.createCell(col)
            cell.setCellValue(COLUMNS[col])
            cell.setCellStyle(headerCellStyle)
        }

        val nbrCellStyle = workbook.createCellStyle()
        nbrCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"))
        val timeCellStyle = workbook.createCellStyle()
        timeCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm:ss.00"))

        var rowIdx = 1
        searchResult.forEachIndexed { idx, result ->
            val row = sheet.createRow(rowIdx++)
            row.createCell(0).setCellValue((idx+1).toDouble())
            row.createCell(1).setCellValue("${participants[result.nbr]?.first}")
            row.createCell(2).setCellValue("${participants[result.nbr]?.second}")
            val pointsCell = row.createCell(3)
            pointsCell.cellStyle = nbrCellStyle
            pointsCell.setCellValue(result.points.toDouble())

            val errorsCell = row.createCell(4)
            errorsCell.cellStyle = nbrCellStyle
            errorsCell.setCellValue(result.errors.toDouble())

            val timeCell = row.createCell(5)
            timeCell.cellStyle = timeCellStyle
            timeCell.setCellValue(Date(result.time.toLong()))

            val sseCell = row.createCell(6)
            if (result.sse) {
                sseCell.setCellValue("x")
            }
        }
    }

    companion object {
        @JvmStatic fun main(args: Array<String>) {
            val fileName = args[0]
            println("Loading data from: ${fileName}")

            val lines: List<String> = File(fileName).readLines()

            val main = Main()
            val loaded = main.loadData(lines)
            if (loaded) {
                main.calcResult()
            }
        }
    }
}

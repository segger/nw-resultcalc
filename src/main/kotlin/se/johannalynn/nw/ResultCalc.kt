package se.johannalynn.nw

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*

class ResultCalc(private val level: Level) {

    enum class Level {
        NW1, NW2
    }

    data class Result(val nbr: Int, val points: Double, val time: Int, val errors: Int, val sse: Int)
    data class TournamentResult(val result: Result, val tp: Double)

    val participants = mutableMapOf<Int, Pair<String,String>>()
    val searchOne = mutableListOf<Result>()
    val searchTwo = mutableListOf<Result>()
    val searchThree = mutableListOf<Result>()
    val searchFour = mutableListOf<Result>()

    val searchNames = mutableListOf<String>()
    val COLUMNS = arrayOf("#","Förare","Hund","Poäng","Fel","Tid","SSE","TP")

    private fun generateResult(nbr: Int, pointsIn: String, errorsIn: String, timeIn: String, sseIn: String): Result {
        val points = if (pointsIn.isNotEmpty()) pointsIn.toDouble() else 0.0
        val errors = if (errorsIn.isNotEmpty()) errorsIn.toInt() else 0
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

        val sse = if (sseIn.isNotEmpty()) 1 else 0

        return Result(nbr, points, time, errors, sse)
    }

    private fun calcTotalResult(resultList: List<List<TournamentResult>>): List<TournamentResult> {
        val results = mutableMapOf<Int, TournamentResult>()

        resultList.forEach { searchList ->
            searchList.forEach { res ->
                val nbr = res.result.nbr
                val curr = results[nbr]
                val newResult = if (curr == null) {
                    res
                } else {
                    val resCopy = res.result.copy(
                            points = curr.result.points + res.result.points,
                            errors = curr.result.errors + res.result.errors,
                            time = curr.result.time + res.result.time,
                            sse = curr.result.sse + res.result.sse
                    )
                    res.copy(result = resCopy, tp = curr.tp + res.tp)
                }
                results[nbr] = newResult
            }
        }

        val comparator = compareByDescending<TournamentResult> { it.result.points }.thenBy { it.result.errors }.thenBy { it.result.time }
        return results.values.sortedWith(comparator)
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
                if (cells.size > 7) {
                    val result2 = generateResult(nbr, cells[7], cells[8], cells[9], cells[10])
                    searchTwo.add(result2)
                }
                if (cells.size > 11) {
                    val result3 = generateResult(nbr, cells[11], cells[12], cells[13], cells[14])
                    searchThree.add(result3)
                }
                if (cells.size > 15) {
                    val result4 = generateResult(nbr, cells[15], cells[16], cells[17], cells[18])
                    searchFour.add(result4)
                }
            }
        }
        return true
    }

    fun calcResultAndPrint(out: String) {
        val workbook = XSSFWorkbook()

        val resultOne = calcResult(searchOne)
        val resultTwo = calcResult(searchTwo)
        val resultThree = calcResult(searchThree)
        val resultFour = calcResult(searchFour)
        val resultTotal = calcTotalResult(listOf(resultOne, resultTwo, resultThree, resultFour))

        printResult(workbook, resultOne, searchNames[0])
        if (resultTwo.isNotEmpty()) {
            printResult(workbook, resultTwo, searchNames[1])
        }
        if (resultThree.isNotEmpty()) {
            printResult(workbook, resultThree, searchNames[2])
        }
        if (resultFour.isNotEmpty()) {
            printResult(workbook, resultFour, searchNames[3])
        }
        printResult(workbook, resultTotal, "Totalen")

        val fileName = "${out}.xlsx"
        println("Write result to: ${fileName}")
        val fileOut = FileOutputStream(fileName)
        workbook.write(fileOut)
        fileOut.close()

        val tournamentFile = "${out}.csv"
        printTournamentResult(tournamentFile, listOf(resultOne, resultTwo, resultThree, resultFour))
    }

    fun calcResult(searchList: List<Result>): List<TournamentResult> {
        val comparator = compareByDescending<Result> { it.points }.thenBy { it.errors }.thenBy { it.time }
        val searchResult = searchList.sortedWith(comparator)

        val tournamentResult = mutableListOf<TournamentResult>()
        searchResult.forEachIndexed { idx, result ->
            val tp = when(level) {
                Level.NW1 -> getTournamentPointsNW1(idx, result)
                else -> getTournamentPointsNW2(idx, result)
            }
            tournamentResult.add(TournamentResult(result, tp))
        }
        return tournamentResult
    }

    fun printResult(workbook: XSSFWorkbook, searchList: List<TournamentResult>, name: String) {
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
        nbrCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#.#"))
        val timeCellStyle = workbook.createCellStyle()
        timeCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm:ss.00"))

        var rowIdx = 1
        searchList.forEachIndexed { idx, tournamentResult ->
            val result = tournamentResult.result
            val row = sheet.createRow(rowIdx++)
            row.createCell(0).setCellValue((idx+1).toDouble())
            row.createCell(1).setCellValue("${participants[result.nbr]?.first}")
            row.createCell(2).setCellValue("${participants[result.nbr]?.second}")
            val pointsCell = row.createCell(3)
            pointsCell.cellStyle = nbrCellStyle
            pointsCell.setCellValue(result.points)

            val errorsCell = row.createCell(4)
            errorsCell.cellStyle = nbrCellStyle
            errorsCell.setCellValue(result.errors.toDouble())

            val timeCell = row.createCell(5)
            timeCell.cellStyle = timeCellStyle
            timeCell.setCellValue(Date(result.time.toLong()))

            val sseCell = row.createCell(6)
            sseCell.cellStyle = nbrCellStyle
            sseCell.setCellValue(result.sse.toDouble())

            val tournamentCell = row.createCell(7)
            tournamentCell.setCellValue(tournamentResult.tp)
        }
    }

    fun printTournamentResult(tournamentFile: String, resultList: List<List<TournamentResult>>) {
        val results = mutableMapOf<Int, MutableList<TournamentResult>>()

        resultList.forEach { searchList ->
            searchList.forEach { res ->
                val nbr = res.result.nbr
                val newResult = results[nbr] ?: mutableListOf()
                newResult.add(res)
                results[nbr] = newResult
            }
        }

        println("Write tournament result to: ${tournamentFile}")
        val tournamentOut = File(tournamentFile)
        val buffer = StringBuffer()
        results.keys.forEach {
            buffer.append("${participants[it]?.first},")
            buffer.append("${participants[it]?.second},")
            results[it]?.forEach {
                buffer.append("${it.tp},")
            }
            buffer.append("\n")
        }
        tournamentOut.writeText(buffer.toString())
    }

    private fun getTournamentPointsNW1(idx: Int, result: Result): Double {
        var tp = 3
        if (result.points > 0) {
            tp += 7;
        }
        val placement = idx + 1
        if (result.points > 0 && placement in 1..5) {
            tp += 6 - placement
        }
        if (result.sse == 1) {
            tp += 5
        }
        return tp.toDouble()
    }

    private fun getTournamentPointsNW2(idx: Int, result: Result): Double {
        var tp = 0.0
        tp += result.points
        val maxPoints = result.points == 25.0
        val placement = idx + 1
        if (maxPoints && placement in 1..5) {
            tp += 6 - placement
        }
        if (result.sse == 1) {
            tp += 5
        }
        tp -= result.errors
        if (tp < 3) {
            return 3.0
        }
        return tp
    }

    companion object {
        @JvmStatic fun main(args: Array<String>) {
            val fileName = args[0]
            println("Loading data from: ${fileName}")

            val lines: List<String> = File(fileName).readLines()

            val main = ResultCalc(Level.valueOf(args[2]))
            val loaded = main.loadData(lines)
            if (loaded) {
                main.calcResultAndPrint(args[1])
            }
        }
    }
}

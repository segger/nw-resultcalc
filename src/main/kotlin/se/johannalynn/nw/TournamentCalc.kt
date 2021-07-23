package se.johannalynn.nw

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*

class TournamentCalc {

    val points = mutableMapOf<Int, List<Double>>()
    val totPoints = mutableMapOf<Int, List<Pair<Int, Double>>>()
    val participants = mutableMapOf<Int, Pair<String,String>>()

    val COLUMNS = arrayOf("#","Förare","Hund","TP")

    data class MaxTournamentResult(val placement: Int, val participant: Int, val maxPoints: Double)

    fun loadData(crNbr: Int, lines: List<String>): Boolean {
        lines.forEach { line ->
            val cells = line.split(",")
            // println(cells)

            val hash = "${cells[0]}${cells[1]}".hashCode()
            participants[hash] = Pair(cells[0], cells[1])

            val pointsList = mutableListOf<Double>()
            var totPoint = Pair(0, 0.0)
            val totCellIdx = cells.size - 1
            cells.forEachIndexed { idx, cell ->
                if (idx in 2 until totCellIdx) {
                    if (cell.isNotEmpty()) {
                        pointsList.add(cell.toDouble())
                    }
                } else if (idx == totCellIdx) {
                    totPoint = Pair(crNbr, cell.toDouble())
                }
            }
            if (points[hash] == null) {
                points[hash] = pointsList
            } else {
                points[hash]?.let { pointsList.addAll(it) }
                points[hash] = pointsList
            }

            if (totPoints[hash] == null) {
                totPoints[hash] = listOf(totPoint)
            } else {
                val tp = mutableListOf(totPoint)
                totPoints[hash]?.let { tp.addAll(it) }
                totPoints[hash] = tp
            }

        }
        return true
    }

    fun calcResult(tournamentFile: String) {
        val workbook = XSSFWorkbook()

        val tournamentResult = calcTopp3TournamentResult()
        val tournamentThreeSearch = calcMaxTournamentResult()
        val tournamentTotResult = calcTournamentTotalResult()
        printResult(workbook, tournamentResult, "Topp 3 CR")
        printResult(workbook, tournamentThreeSearch, "Topp 3 sök")
        printResult(workbook, tournamentTotResult, "Total")

        println("Write result to: ${tournamentFile}")
        val fileOut = FileOutputStream(tournamentFile)
        workbook.write(fileOut)
        fileOut.close()
    }

    private fun calcMaxTournamentResult(): List<MaxTournamentResult> {
        val maxPoints = mutableMapOf<Int, Double>()
        points.forEach {
            val pointList = it.value.sortedDescending()
            val maxPoint = pointList[0] + pointList[1] + pointList[2]
            maxPoints[it.key] = maxPoint
        }

        val sorted = maxPoints.toList().sortedByDescending { (_, v) -> v }
        return sorted.mapIndexed { idx, res ->
            MaxTournamentResult(idx+1, res.first, res.second)
        }
    }

    private fun calcTopp3TournamentResult(): List<MaxTournamentResult> {
        val topp3Points = mutableMapOf<Int, Double>()
        totPoints.forEach {
            val tp = it.value.sortedByDescending { it.second }.take(3)
            val sum = tp.map { it.second }.sum()
            topp3Points[it.key] = sum
        }

        val sorted = topp3Points.toList().sortedByDescending { (_, v) -> v }
        return sorted.mapIndexed { idx, res ->
            MaxTournamentResult(idx + 1, res.first, res.second)
        }
    }

    private fun calcTournamentTotalResult(): List<MaxTournamentResult> {
        val totalPoints = mutableMapOf<Int, Double>()
        points.forEach {
            val pointList = it.value.sortedDescending()
            val totalPoint = pointList.sum()
            totalPoints[it.key] = totalPoint
        }

        val sorted = totalPoints.toList().sortedByDescending { (_, v) -> v }
        return sorted.mapIndexed { idx, res ->
            MaxTournamentResult(idx+1, res.first, res.second)
        }
    }

    fun printResult(workbook: XSSFWorkbook, resultList: List<MaxTournamentResult>, name: String) {
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

        var rowIdx = 1
        resultList.forEach { result ->
            val row = sheet.createRow(rowIdx++)
            row.createCell(0).setCellValue(result.placement.toDouble())
            row.createCell(1).setCellValue("${participants[result.participant]?.first}")
            row.createCell(2).setCellValue("${participants[result.participant]?.second}")
            val pointsCell = row.createCell(3)
            pointsCell.cellStyle = nbrCellStyle
            pointsCell.setCellValue(result.maxPoints.toDouble())
        }
    }

    companion object {
        @JvmStatic fun main(args: Array<String>) {
            val level = args[0]
            val dirName = args[1]
            println("Loading data from: ${dirName}")

            val tournamentCsvFiles = mutableListOf<File>()
            File(dirName).listFiles().forEach {
                val isCsvFile = it.name.endsWith(".csv") && it.name.startsWith(level)
                if (isCsvFile) {
                    tournamentCsvFiles.add(it)
                }
            }

            val main = TournamentCalc()
            tournamentCsvFiles.forEach {
                val tmp = it.name.substringBefore(".")
                val nbr = tmp.substring(tmp.length - 1).toInt()
                main.loadData(nbr, it.readLines())
            }

            main.calcResult(args[2])
        }
    }
}

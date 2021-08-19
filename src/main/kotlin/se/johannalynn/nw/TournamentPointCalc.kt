package se.johannalynn.nw

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*

class TournamentPointCalc(private val level: Level) {

    enum class Level {
        NW1, NW2
    }

    val points = mutableMapOf<Int, List<Pair<Int, MutableList<Double>>>>()
    val totPoints = mutableMapOf<Int, List<Pair<Int, Double>>>()
    val stroked = mutableMapOf<Int, Int>()
    val participants = mutableMapOf<Int, Pair<String,String>>()

    val COLUMNS = arrayOf("#","Förare","Hund","TP", "Struken CR")
    val CR_COLUMNS = mutableListOf<Pair<Int, String>>()
    data class MaxTournamentResult(val placement: Int, val participant: Int, val maxPoints: Double)

    fun loadData(crNbr: Int, lines: List<String>): Boolean {
        // println("FILE $crNbr")
        lines.forEachIndexed { index, line ->
            if (index == 0) {
                val momentCells = line.split(",")
                momentCells.forEach { name ->
                    CR_COLUMNS.add(Pair(crNbr, name))
                }
                CR_COLUMNS.add(Pair(crNbr, "Tot TP"))
            } else {
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
                            val points = cell.toDouble()
                            pointsList.add(points)
                        }
                    } else if (idx == totCellIdx) {
                        totPoint = Pair(crNbr, cell.toDouble())
                    }
                }
                val crPointsList = Pair(crNbr, pointsList)

                if (points[hash] == null) {
                    points[hash] = listOf(crPointsList)
                } else {
                    val crPoints = mutableListOf(crPointsList)
                    points[hash]?.let { crPoints.addAll(it) }
                    points[hash] = crPoints
                }

                if (totPoints[hash] == null) {
                    totPoints[hash] = listOf(totPoint)
                } else {
                    val tp = mutableListOf(totPoint)
                    totPoints[hash]?.let { tp.addAll(it) }
                    totPoints[hash] = tp
                }
            }
        }
        return true
    }

    fun calcResult(tournamentFile: String) {
        val workbook = XSSFWorkbook()

        val tournamentResult = calcTopp3TournamentResult()
        printResult(workbook, tournamentResult, "Topp 3 CR")

        println("Write result to: ${tournamentFile}")
        val fileOut = FileOutputStream(tournamentFile)
        workbook.write(fileOut)
        fileOut.close()
    }

    private fun calcTopp3TournamentResult(): List<MaxTournamentResult> {
        val topp3Points = mutableMapOf<Int, Double>()
        totPoints.forEach {
            val sorted = it.value.sortedByDescending { it.second }
            val tp = sorted.take(3)
            val last = sorted.takeLast(1).first().first
            val sum = tp.map { it.second }.sum()
            topp3Points[it.key] = sum
            if (sorted.size == 4) {
                stroked[it.key] = last
            }
        }

        val sorted = topp3Points.toList().sortedByDescending { (_, v) -> v }
        return sorted.mapIndexed { idx, res ->
            MaxTournamentResult(idx + 1, res.first, res.second)
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
        var colIdx = 0
        for (col in COLUMNS.indices) {
            val cell = headerRow.createCell(colIdx++)
            cell.setCellValue(COLUMNS[col])
            cell.setCellStyle(headerCellStyle)
        }
        for (crCol in CR_COLUMNS.indices) {
            val cell = headerRow.createCell(colIdx++)
            cell.setCellValue(CR_COLUMNS[crCol].second)
            // cell.setCellStyle(headerCellStyle)
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

            val crCell = row.createCell(4)
            val participantStroke = stroked[result.participant]
            crCell.setCellValue(participantStroke?.toString())

            val participantPoints = points[result.participant]
            participantPoints?.sortedBy { it.first }?.forEach {
                var pColIdx = colIdx(it.first)
                var sum = 0.0
                it.second.forEach {
                    val pCell = row.createCell(pColIdx++)
                    pCell.cellStyle = nbrCellStyle
                    pCell.setCellValue(it)
                    sum += it
                }
                val totCell = row.createCell(pColIdx++)
                totCell.cellStyle = headerCellStyle
                totCell.setCellValue(sum)
            }
        }
    }

    private fun colIdx(order: Int): Int {
        return when(order) {
            1 -> 5
            2 -> if (level == Level.NW1) 10 else 9
            3 -> if (level == Level.NW1) 15 else 13
            4 -> if (level == Level.NW1) 20 else 17
            else -> 0
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
            val main = TournamentPointCalc(Level.valueOf(level))
            tournamentCsvFiles.sorted().forEach {
                val tmp = it.name.substringBefore(".")
                val nbr = tmp.substring(tmp.length - 1).toInt()
                main.loadData(nbr, it.readLines())
            }

            main.calcResult(args[2])
        }
    }
}

package se.johannalynn.nw

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*

class TournamentPoints(private val level: Level) {

    enum class Level {
        NW1, NW2
    }

    val points = mutableMapOf<Int, List<Pair<Int, Double>>>()
    val totPoints = mutableMapOf<Int, List<Pair<Int, Double>>>()
    val participants = mutableMapOf<Int, Pair<String,String>>()
    val stroked = mutableMapOf<Int, List<Int>>()

    val COLUMNS = arrayOf("#","Förare","Hund","TP","Strukna")
    val CR_COLUMNS = mutableListOf<Pair<Int, String>>()

    data class MaxTournamentResult(val placement: Int, val participant: Int, val maxPoints: Double)

    private fun nbrOfSearch(): Int {
        return if (level == Level.NW1) 4 else 3
    }

    fun loadData(crNbr: Int, lines: List<String>): Boolean {
        // println("FILE $crNbr")
        lines.forEachIndexed { index, line ->
            if (index == 0) {
                val momentCells = line.split(",")
                momentCells.forEachIndexed { idx, name ->
                    val nbr = (crNbr-1)*(nbrOfSearch()) + idx
                    CR_COLUMNS.add(Pair(nbr, "$name (${nbr + 1})"))
                }
            } else {
                val cells = line.split(",")
                // println(cells)

                val hash = "${cells[0]}${cells[1]}".hashCode()
                participants[hash] = Pair(cells[0], cells[1])

                val pointsList = mutableListOf<Pair<Int, Double>>()
                var totPoint = Pair(0, 0.0)
                val totCellIdx = cells.size - 1
                cells.forEachIndexed { idx, cell ->
                    if (idx in 2 until totCellIdx) {
                        if (cell.isNotEmpty()) {
                            val points = cell.toDouble()
                            pointsList.add(Pair((crNbr-1)*(nbrOfSearch())+idx, points))
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

        }
        return true
    }

    fun calcResult(tournamentFile: String) {
        val workbook = XSSFWorkbook()

        val tournamentThreePartResult = calcMaxThreeTournametResult()
        val sheetName = if (level == Level.NW1) "12" else "9"
        printResult(workbook, tournamentThreePartResult, "Topp $sheetName sök")

        println("Write result to: ${tournamentFile}")
        val fileOut = FileOutputStream(tournamentFile)
        workbook.write(fileOut)
        fileOut.close()
    }

    private fun take(): Int {
        if (level == Level.NW1) {
            return 4*3
        } else if (level == Level.NW2) {
            return 3*3
        }
        return 4*3
    }

    private fun calcMaxThreeTournametResult(): List<MaxTournamentResult> {
        val maxPoints = mutableMapOf<Int, Double>()
        points.forEach {
            val pointList = it.value.sortedByDescending{ it.second }
            val maxPointTakes = pointList.take(take())
            val maxPoint = maxPointTakes.map { it.second }.sum()
            maxPoints[it.key] = maxPoint
            if (pointList.size > take()) {
                val lastTake = pointList.takeLast(pointList.size - take())
                stroked[it.key] = lastTake.sortedBy { it.first }.map { it.first - 1 }
            }
        }

        val sorted = maxPoints.toList().sortedByDescending { (_, v) -> v }
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

        val strikeFont = workbook.createFont()
        strikeFont.setStrikeout(true)
        val strikeThroughStyle = workbook.createCellStyle()
        strikeThroughStyle.setFont(strikeFont)
        strikeThroughStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"))

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
                val pColIdx = 3 + it.first
                val pCell = row.createCell(pColIdx)
                pCell.cellStyle = nbrCellStyle
                if (participantStroke != null && participantStroke.contains(it.first - 1)) {
                    pCell.cellStyle = strikeThroughStyle
                }
                pCell.setCellValue(it.second)
            }
        }
    }

    private fun colIdx(order: Int): Int {
        return when(order) {
            1 -> 5
            2 -> if (level == Level.NW1) 9 else 8
            3 -> if (level == Level.NW1) 13 else 11
            4 -> if (level == Level.NW1) 17 else 14
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

            val main = TournamentPoints(Level.valueOf(level))
            tournamentCsvFiles.sorted().forEach {
                val tmp = it.name.substringBefore(".")
                val nbr = tmp.substring(tmp.length - 1).toInt()
                main.loadData(nbr, it.readLines())
            }

            main.calcResult(args[2])
        }
    }
}

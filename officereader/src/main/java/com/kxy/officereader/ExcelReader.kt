package com.kxy.officereader

import android.util.Log
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Color
import org.apache.poi.ss.usermodel.Drawing
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFDrawing
import org.apache.poi.xssf.usermodel.XSSFPicture
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.IOException

class ExcelReader {
    // 读取 Excel 文件
    public fun readExcel(filePath: String): StringBuilder {
        val fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length)
        if (fileType != "xlsx") {
            return  java.lang.StringBuilder("格式错误")
        }

        val htmlContent = StringBuilder()
        try {
            val workbook = XSSFWorkbook(filePath)

            // 2. 生成HTML代码
            htmlContent.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/html401'>")
            htmlContent.append("<head>")
            htmlContent.append("<meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>")
            htmlContent.append("</head><body><div class=\"container\">")
            // 3. 遍历每个工作表
            for (sheetIndex in 0 until workbook.numberOfSheets) {
                val sheet: Sheet = workbook.getSheetAt(sheetIndex)
                htmlContent.append("<a class=\"btn btn-block btn-default\" style=\"margin:10px 0;font-size:large;text-align:left;overflow:hidden;text-overflow:ellipsis;\" data-toggle=\"collapse\" href=\"#t_468047691${sheetIndex}\" aria-expanded=\"true\"><b>${sheet.sheetName}</b></a>")
                htmlContent.append("<div id=\"t_468047691${sheetIndex}\" class=\"collapse in\"><div class=\"table-responsive table-sheet-wrap\">")
                htmlContent.append("<table class=\"table table-bordered table-hover\" style=\" width:946px;border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">")
                // 4. 遍历每一行
                for (row in sheet) {
                    val height = row.height / 15.625
                    htmlContent.append("<tr height=\"$height\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">")
                    // 5. 遍历每一列
                    for (cell in row) {
                        if (cell.cellType == CellType.BLANK) {
                            continue
                        } else {
                            val tdStyle =
                                StringBuilder("<td style=\"border:1px solid #000; border-width:0 1px 1px 0;margin:2px 0 2px 0; ")
                            val cellStyle = cell.cellStyle
                            if (cellStyle.fillForegroundColorColor != null) {
                                val bgColor = getCellHexColor(cellStyle.fillForegroundColorColor)
                                if (!bgColor.isNullOrEmpty()) {
                                    tdStyle.append(" background-color:$bgColor;")
                                }

                            }

                            htmlContent.append(tdStyle)

                            val width = (sheet.getColumnWidth(cell.columnIndex) / 35.7)
                            val cellReginCol =
                                getMergerCellRegionCol(sheet, cell.rowIndex, cell.columnIndex)
                            val cellReginRow =
                                getMergerCellRegionRow(sheet, cell.rowIndex, cell.columnIndex)
                            val hAlign = convertAlignToHtml(cellStyle.alignment)
                            val vAlign = convertVerticalAlignToHtml(cellStyle.verticalAlignment)

                            htmlContent.append(" align=\"$hAlign\"")
                            htmlContent.append(" valign=\"$vAlign\"")
                            htmlContent.append(" width=\"$width\"")
                            if (cellReginCol > 1) {
                                htmlContent.append(" colspan=\"$cellReginCol\"")
                            }
                            if (cellReginRow > 1) {
                                htmlContent.append(" rowspan=\"$cellReginRow\"")
                            }

                            htmlContent.append(">")

                            extractedTextContent(htmlContent, hAlign, workbook, cellStyle, cell)

                            htmlContent.append("</td>")

                        }
                    }

                    htmlContent.append("</tr>")
                }
                //读取图片
                if (sheet.drawingPatriarch != null) {
                    getPictures(sheet.drawingPatriarch, htmlContent, sheet)
                }

                htmlContent.append("</table></div></div>")
            }

            htmlContent.append("</div></body></html>")

        } catch (e: IOException) {
            e.printStackTrace()
        }
        return htmlContent

    }

    private fun extractedTextContent(
        htmlContent: StringBuilder,
        hAlign: String?,
        workbook: XSSFWorkbook,
        cellStyle: CellStyle,
        cell: Cell
    ) {
        htmlContent.append("<p style=\"text-align:$hAlign;\">")

        if (cell.cellType == CellType.STRING && cell.richStringCellValue.numFormattingRuns() > 0) {
            // 4. 获取单元格内容
            val richTextString = cell.richStringCellValue as XSSFRichTextString
            // 5. 遍历富文本内容
            for (i in 0 until richTextString.numFormattingRuns()) {
                val startIndex = richTextString.getIndexOfFormattingRun(i)
                val endIndex = if (i < richTextString.numFormattingRuns() - 1) {
                    richTextString.getIndexOfFormattingRun(i + 1)
                } else {
                    richTextString.length()
                }
                // 获取文本
                var text = richTextString.string.substring(startIndex, endIndex)
                // 获取字体信息
                val font = richTextString.getFontAtIndex(startIndex)

                // 获取字体大小
                val fontSize = font.fontHeightInPoints
                for (index in 0 until countNewlines(text)) {
                    text = text.replace("\n", "<br>")
                }


                //获取下划线信息
                var underlineH5String = when (font.underline) {
                    Font.U_SINGLE -> "text-decoration: underline"
                    Font.U_DOUBLE -> "text-decoration: underline double"
                    Font.U_SINGLE_ACCOUNTING -> "text-decoration: underline wavy"
                    Font.U_DOUBLE_ACCOUNTING -> "text-decoration: underline wavy double"
                    else -> ""
                }

                underlineH5String += if (font.strikeout) " line-through" else ";"


                // 获取字体颜色
                val fontColor = font.xssfColor
                val fontColorHex = getCellHexColor(fontColor) ?: "#000000"
                htmlContent.append("<span  style=\"")
                htmlContent.append("font-size:$fontSize;")
                htmlContent.append("font-family: ${font.fontName};")
                if (font.bold) htmlContent.append(" font-weight:bold;")
                htmlContent.append(" $underlineH5String")
                htmlContent.append(" color:$fontColorHex;\">")
                if (font.italic) htmlContent.append("<i>")
                htmlContent.append(text)
                if (font.italic) htmlContent.append("</i>")
                htmlContent.append("</span>")
            }

        } else {
            var text = cellToString(cell)
            for (index in 0 until countNewlines(text)) {
                text = text.replace("\n", "<br>")
            }

            htmlContent.append("<span  style=\"")
            val font = workbook.getFontAt(cellStyle.fontIndex)
            val fontColor = getCellHexColor(font.xssfColor)
            val isBold = font.bold
            val fontHeight = font.fontHeightInPoints

            if (!fontColor.isNullOrEmpty()) {
                htmlContent.append(" color:$fontColor;")
            }

            if (isBold) {
                htmlContent.append(" font-weight:bold;")
            }
            htmlContent.append(" font-size:$fontHeight;")
            htmlContent.append("font-family: ${font.fontName};\"")
            htmlContent.append(">")
            if (font.italic) htmlContent.append("<i>")
            htmlContent.append(text)
            if (font.italic) htmlContent.append("</i>")
            htmlContent.append("</span>")
        }



        htmlContent.append("</p>")

    }

    private fun cellToString(cell: Cell): String {
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> cell.numericCellValue.toString()
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            else -> ""
        }
    }

    private fun getCellHexColor(color: Color?): String? {
        if (color is XSSFColor) {
            if (color.isRGB) {
                val rgb = color.argb
                if (rgb != null) {
                    return String.format("#%02X%02X%02X", rgb[1], rgb[2], rgb[3])
                }
            } else {
                return null
            }

        }
        return null
    }

    @Throws(IOException::class)
    private fun getMergerCellRegionCol(
        sheet: Sheet, cellRow: Int,
        cellCol: Int
    ): Int {
        var retVal = 0
        val sheetMergerCount = sheet.numMergedRegions
        for (i in 0 until sheetMergerCount) {
            val cra = sheet.getMergedRegion(i) as CellRangeAddress
            val firstRow = cra.firstRow
            val firstCol = cra.firstColumn
            val lastRow = cra.lastRow
            val lastCol = cra.lastColumn
            if (cellRow in firstRow..lastRow) {
                if (cellCol in firstCol..lastCol) {
                    retVal = lastCol - firstCol + 1
                    break
                }
            }
        }
        return retVal
    }

    @Throws(IOException::class)
    private fun getMergerCellRegionRow(
        sheet: Sheet, cellRow: Int,
        cellCol: Int
    ): Int {
        var retVal = 0
        val sheetMergerCount = sheet.numMergedRegions
        for (i in 0 until sheetMergerCount) {
            val cra = sheet.getMergedRegion(i) as CellRangeAddress
            val firstRow = cra.firstRow
            val firstCol = cra.firstColumn
            val lastRow = cra.lastRow
            val lastCol = cra.lastColumn
            if (cellRow in firstRow..lastRow) {
                if (cellCol in firstCol..lastCol) {
                    retVal = lastRow - firstRow + 1
                    break
                }
            }
        }
        return retVal
    }

    private fun convertAlignToHtml(alignment: HorizontalAlignment): String? {
        var align = "left"
        when (alignment) {
            HorizontalAlignment.LEFT -> align = "left"
            HorizontalAlignment.CENTER -> align = "center"
            HorizontalAlignment.RIGHT -> align = "right"
            else -> {}
        }
        return align
    }

    private fun convertVerticalAlignToHtml(verticalAlignment: VerticalAlignment): String? {
        var valign = "middle"
        when (verticalAlignment) {
            VerticalAlignment.BOTTOM -> valign = "bottom"
            VerticalAlignment.CENTER -> valign = "center"
            VerticalAlignment.TOP -> valign = "top"
            else -> {}
        }
        return valign
    }

    private fun countNewlines(inputString: String): Int {
        val regex = Regex("\n")
        val matches = regex.findAll(inputString)
        return matches.count()
    }

    private fun getPictures(drawing: Drawing<*>, htmlContent: StringBuilder, sheet: Sheet) {
        if (drawing is XSSFDrawing) {
            val xssfDrawing = drawing as XSSFDrawing
            xssfDrawing.shapes.forEach { xssfShape ->
                if (xssfShape is XSSFPicture) {
                    val imageData = xssfShape.pictureData.data
                    val base64Image = java.util.Base64.getEncoder().encodeToString(imageData)

                    var left = 0.0
                    for (i in 0 until xssfShape.clientAnchor.col1) {
                        left += (sheet.getColumnWidth(i).toFloat() / 35.7)
                    }

                    var top = 0.0
                    for (i in 0 until xssfShape.clientAnchor.row1) {
                        top += (sheet.getRow(i).height.toFloat() / 15.625)
                    }
                    top += xssfShape.clientAnchor.dy1 / 91440 * 32
                    htmlContent.append(
                        "<div class=\"image-container\">\n" +
                                "  <img style=\"position: absolute; left: ${left}; top:${top}\" class=\"excel-image\" src=\"data:image/png;base64," + base64Image + "\" alt=\"Excel Image\">\n" +
                                "</div>\n"
                    )
                }
            }
        }
    }
}


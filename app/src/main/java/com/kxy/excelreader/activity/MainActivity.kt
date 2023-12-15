package com.kxy.excelreader.activity

import android.Manifest
import android.annotation.SuppressLint
import android.content.pm.PackageManager
import android.os.Bundle
import android.util.Log
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.kxy.excelreader.R
import com.kxy.excelreader.databinding.ActivityMainBinding
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Color
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import java.io.File
import java.io.IOException


class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(ActivityMainBinding.inflate(layoutInflater).root)
        readExcel(getFile())
//        val intent = Intent(this, ExcelRead::class.java)
//        intent.putExtra("name", "/storage/emulated/0/Download/四川省工程技术人员职称申报评审基本条件.xls")
//        startActivity(intent)
    }

    private fun getFile(): String {
        if (ContextCompat.checkSelfPermission(
                this,
                Manifest.permission.WRITE_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            // 如果没有权限，则请求权限
            val REQUEST_CODE_STORAGE_PERMISSION = 10086
            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),
                REQUEST_CODE_STORAGE_PERMISSION
            )
        }
        val file: File = File(
            "/storage/emulated/0/Download/四川省工程技术人员职称申报评审基本条件.xlsx"
        )

        return "/storage/emulated/0/Download/四川省工程技术人员职称申报评审基本条件.xlsx"
    }

    // 读取 Excel 文件
    @SuppressLint("SetJavaScriptEnabled")
    private fun readExcel(filePath: String?) {
        try {
            val workbook = XSSFWorkbook(filePath)
            // 2. 生成HTML代码
            val htmlContent = StringBuilder()
            htmlContent.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/html401'>")
            htmlContent.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet></head>")

            // 3. 遍历每个工作表
            for (sheetIndex in 0 until workbook.numberOfSheets) {
                val sheet: Sheet = workbook.getSheetAt(sheetIndex)
                htmlContent.append("<table style=\" width:946px;border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">")
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
            }

            htmlContent.append("</table></body></html>")

            val webView = findViewById<WebView>(R.id.webView)
            val setting = webView.settings
            setting.javaScriptEnabled = true
            setting.builtInZoomControls = true
            setting.cacheMode = WebSettings.LOAD_CACHE_ELSE_NETWORK
//            webView.setInitialScale(300)
            Log.e("HTML内容", htmlContent.toString())
            webView.loadData(htmlContent.toString(), "text/html", "utf-8")
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    private fun extractedTextContent(
        htmlContent: StringBuilder,
        hAlign: String?,
        workbook: XSSFWorkbook,
        cellStyle: CellStyle,
        cell: Cell
    ) {
        htmlContent.append("<p style=\"text-align:$hAlign;\">")

        if (cell.cellType == CellType.STRING) {
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
                val text = richTextString.string.substring(startIndex, endIndex)
                // 获取字体信息
                val fontIndex = richTextString.getIndexOfFormattingRun(i)

                val font = richTextString.getFontAtIndex(fontIndex)

                // 获取字体大小
                val fontSize = font.fontHeightInPoints

                // 获取字体颜色
                val fontColor = font.xssfColor
                println(" 富文本内容 fontColor: $fontColor")
                val fontColorHex = getCellHexColor(fontColor) ?: "#000000"

                // 输出信息
                println(" 富文本内容 Text: $text")
                println(" 富文本内容 startIndex: $startIndex")
                println(" 富文本内容 endIndex: $endIndex")
                println(" 富文本内容 fontIndex: $fontIndex")
                println(" 富文本内容 Font Size: $fontSize")
                println(" 富文本内容 Font Color: $fontColorHex")
                htmlContent.append("<span  style=\"")
                htmlContent.append("font-size:$fontSize;")
                htmlContent.append(" color:$fontColorHex;\">")
                htmlContent.append(text)
                htmlContent.append("</span>")

            }
        } else {
            htmlContent.append("<span  style=\"")
            val font = workbook.getFontAt(cellStyle.fontIndex)
            val fontColor = getCellHexColor(font.xssfColor)
            val isBold = font.bold
            val fontHeight = font.fontHeight / 2

            if (!fontColor.isNullOrEmpty()) {
                htmlContent.append(" color:$fontColor;")
            }

            if (isBold) {
                htmlContent.append(" font-weight:bold;")
            }


            htmlContent.append(" font-size:$fontHeight%;\"")
            htmlContent.append(">" + cellToString(cell) + "</span>")
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
            val rgb = color.argb
            if (rgb != null) {
                return String.format("#%02X%02X%02X", rgb[1], rgb[2], rgb[3])
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
            val firstRow = cra.firstRow // �ϲ���Ԫ��CELL��ʼ��
            val firstCol = cra.firstColumn // �ϲ���Ԫ��CELL��ʼ��
            val lastRow = cra.lastRow // �ϲ���Ԫ��CELL������
            val lastCol = cra.lastColumn // �ϲ���Ԫ��CELL������
            if (cellRow in firstRow..lastRow) { // �жϸõ�Ԫ���Ƿ����ںϲ���Ԫ����
                if (cellCol in firstCol..lastCol) {
                    retVal = lastRow - firstRow + 1 // �õ��ϲ�������
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
}


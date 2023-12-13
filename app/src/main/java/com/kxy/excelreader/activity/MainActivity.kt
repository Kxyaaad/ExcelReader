package com.kxy.excelreader.activity

//import com.alibaba.excel.EasyExcel
//import com.alibaba.excel.context.AnalysisContext
//import com.alibaba.excel.enums.CellExtraTypeEnum
//import com.alibaba.excel.event.AnalysisEventListener
//import com.alibaba.excel.metadata.CellExtra
//import com.alibaba.excel.support.ExcelTypeEnum
//import com.alibaba.excel.util.ListUtils
import android.Manifest
import android.content.pm.PackageManager
import android.os.Bundle
import android.util.Log
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.kxy.excelreader.databinding.ActivityMainBinding
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(ActivityMainBinding.inflate(layoutInflater).root)
        readExcel(getFile())
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

        Log.e(TAG, "是否是文件" + file.isFile.toString())
        return "/storage/emulated/0/Download/四川省工程技术人员职称申报评审基本条件.xlsx"
    }

    // 读取 Excel 文件
    private fun readExcel(filePath: String?) {
//        var excelReader: ExcelReader? = null
//        try {
//            excelReader = EasyExcel.read(filePath).build()
//            val readSheet = EasyExcel.readSheet(0).build()
//            excelReader.read(readSheet)
//        } finally {
//            excelReader?.finish()
//        }
//        EasyExcel.read(filePath, NoModelDataListener())
//            .extraRead(CellExtraTypeEnum.COMMENT)
//            .extraRead(CellExtraTypeEnum.MERGE)
//            .extraRead(CellExtraTypeEnum.MERGE)
//            .excelType(ExcelTypeEnum.XLSX)
//            .sheet().doRead()
//
//    }
        val workbook = XSSFWorkbook(filePath)
        val sheet = workbook.getSheetAt(0)
        for (row in sheet) {
            for (cell:Cell in row) {
                Log.e(TAG, "cell type  ${cell.cellType}")
                Log.e(TAG, "richStringCellValue  ${cell.richStringCellValue}")

                val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                Log.e(TAG, "Font Color:  ${font.boldweight}")
                Log.e(TAG, "Font Size:  ${font.fontHeightInPoints}")

                val cellRef = CellReference(cell)
                Log.e(TAG, "row ${cellRef.row}")
                Log.e(TAG, "col ${cellRef.col}")
                Log.e(TAG, "背景色 ${cell.cellStyle.fillForegroundColorColor}")
            }
        }

        // 3. 获取所有合并单元格区域
        // 3. 获取所有合并单元格区域
       Log.e(TAG, "合并单元格 ${sheet.numMergedRegions}")
        val mergedRegion = sheet.getMergedRegion(50)
        // 4. 检查是否有合并单元格

        // 4. 检查是否有合并单元格
//        if (mergedRegions.size > 0) {
//            println("Sheet contains merged cells.")
//
//            // 5. 遍历合并单元格区域
//            for (mergedRegion in mergedRegions) {
                val firstRow: Int = mergedRegion.getFirstRow()
                val lastRow: Int = mergedRegion.getLastRow()
                val firstCol: Int = mergedRegion.getFirstColumn()
                val lastCol: Int = mergedRegion.getLastColumn()
                Log.e(TAG, "Merged Region: ($firstRow, $firstCol) to ($lastRow, $lastCol)")
//            }
//        } else {
//            println("Sheet does not contain merged cells.")
//        }
    }


    val TAG = "excelReader"
//
//class NoModelDataListener : AnalysisEventListener<Map<Int, String>>() {
//
//    /**
//     * 每隔5条存储数据库，实际使用中可以100条，然后清理list ，方便内存回收
//     */
//    private val BATCH_COUNT = 5
//    private var cachedDataList = ListUtils.newArrayListWithExpectedSize<Map<Int, String>>(BATCH_COUNT)
//
//    override fun invoke(data: Map<Int, String>, context: AnalysisContext) {
//        Log.e(TAG, data.entries.toString())
//        Log.e("$TAG 解析到一条数据:{}", Gson().toJson(data).toString())
//
//
//        cachedDataList.add(data)
//        if (cachedDataList.size >= BATCH_COUNT) {
//            //每5条加载一次数据
//            saveData()
//            cachedDataList = ListUtils.newArrayListWithExpectedSize<Map<Int, String>>(BATCH_COUNT)
//        }
//    }
//
//
//    override fun extra(extra: CellExtra, context: AnalysisContext?) {
//        Log.e("读取到了一条额外信息:{}", Gson().toJson(extra).toString())
//        when (extra.type) {
//            CellExtraTypeEnum.COMMENT -> Log.e(
//                "额外信息是批注,在rowIndex:{},columnIndex;{},内容是:{}", extra.rowIndex.toString() + "=>" +
//                extra.columnIndex + "=>" +
//                extra.text
//            )
//
//            CellExtraTypeEnum.HYPERLINK -> if ("Sheet1!A1" == extra.text) {
//                Log.e(
//                    "额外信息是超链接,在rowIndex:{},columnIndex;{},内容是:{}", extra.rowIndex.toString() + "=>" +
//                    extra.columnIndex + "=>" + extra.text
//                )
//            } else if ("Sheet2!A1" == extra.text) {
//                Log.e(
//                    "额外信息是超链接,而且覆盖了一个区间,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{},"
//                            + "内容是:{}",
//                    extra.firstRowIndex.toString()+ "=>" + extra.firstColumnIndex + "=>" + extra.lastRowIndex + "=>" +
//                    extra.lastColumnIndex + "=>" + extra.text
//                )
//            } else {
////                Assertions.fail("Unknown hyperlink!")
//            }
//
//            CellExtraTypeEnum.MERGE -> Log.e(
//                "额外信息是合并单元格,而且覆盖了一个区间,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{}",
//                extra.firstRowIndex.toString() + "=>" + extra.firstColumnIndex + "=>" + extra.lastRowIndex + "=>" +
//                extra.lastColumnIndex
//            )
//
//            else -> {}
//        }
//    }
//
//    override fun doAfterAllAnalysed(context: AnalysisContext) {
//        saveData()
//        Log.e(TAG,"所有数据解析完成！")
//    }
//
//    /**
//     * 加上存储数据库
//     */
//    private fun saveData() {
//        Log.e(TAG, "{}条数据，开始存储数据库！${cachedDataList.size}")
//        // 存储数据库成功的逻辑
//        Log.e(TAG,"存储数据库成功！")
//    }
}


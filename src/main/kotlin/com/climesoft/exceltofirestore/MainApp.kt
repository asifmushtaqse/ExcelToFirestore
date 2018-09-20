package com.climesoft.exceltofirestore

import com.google.firebase.FirebaseApp
import com.google.auth.oauth2.GoogleCredentials
import com.google.firebase.FirebaseOptions
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import kotlin.math.round
import com.google.firebase.cloud.FirestoreClient


fun main(args: Array<String>){
    initFirestore();
    readExcel();
}

fun readExcel(){
    val excelFile = FileInputStream("src/main/resources/data.xlsx")
    val xlWb = WorkbookFactory.create(excelFile)
    val array = ArrayList<HashMap<String, Any>>()
    val xlWs = xlWb.getSheetAt(0)
    var skip = true
    for(row in xlWs.rowIterator()){
        if(!skip){
            val dn = row.getCell(0).stringCellValue
            val id = row.getCell(1).stringCellValue
            val cell = row.getCell(2)
            var s: Any = ""
            if(cell.cellType === CellType.FORMULA)
                s = when (cell.cachedFormulaResultType) {
                    CellType.NUMERIC -> round(cell.numericCellValue).toLong()
                    CellType.STRING -> cell.richStringCellValue.string
                    else -> ""
                }
            val tc = round(row.getCell(3).numericCellValue).toLong()
            val map = HashMap<String, Any>()
            map.put("dn", dn)
            map.put("id", id)
            map.put("s", s)
            map.put("tc", tc)
            array.add(map)
        }
        skip = false
    }
    val data = HashMap<String, Any>()
    data.put("data", array)
    val db = FirestoreClient.getFirestore()
    println(db.collection("portfolioDetails").add(data).get().id)

}
fun initFirestore(){
    val serviceAccount = FileInputStream("src/exceltofirestore-7222c-firebase-adminsdk-1s8nt-c7984f0061.json")
    val options = FirebaseOptions.Builder()
            .setCredentials(GoogleCredentials.fromStream(serviceAccount))
            .setDatabaseUrl("https://exceltofirestore-7222c.firebaseio.com")
            .build()
    FirebaseApp.initializeApp(options)
}
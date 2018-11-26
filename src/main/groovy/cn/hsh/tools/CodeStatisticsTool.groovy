package cn.hsh.tools
/**
 * author：HSH
 * time:10:44
 **/
class CodeStatisticsTool {
    static int    MAX_LINE_SIZE = 500
    static String FILE_NAME_KW  = "groovy"

    static main(def args) {
        String filePath = "E:\\workspace\\lh-eia-corp-eia-web"
        def content = []
        def statisticsLineSize = [500,1000]
        statisticsLineSize.each {
            def sheetNode = [:]
            sheetNode.sheetName = it+"行"
            sheetNode.title = ["文件名", "代码行数", "文件路径"]
            sheetNode.content = []
            sheetNode.lineSize = it
            content << sheetNode
        }
        List<File>  fileList = searchFileListByLineSizeAndKW(new File(filePath), FILE_NAME_KW)
        fileList.each {
            def lineSize = it.readLines().size()
            for(def sheetNode: content){
                if( lineSize >= sheetNode.lineSize){
                    def contentNode = []
                    contentNode << it.name
                    contentNode << lineSize
                    contentNode << it.path
                    sheetNode.content << contentNode
                }
            }
        }
        ExportExcelTools.exportGeoExcel(content,filePath+"\\代码统计.xls")
    }

    static List<File> searchFileListByLineSizeAndKW(File folder, String keywords) {
        def resultList = []
        File[] subFiles = folder.listFiles(new FileFilter() {
            @Override
            boolean accept(File sub) {
                if (sub.isDirectory()) {
                    return true
                } else if (sub.getName().toLowerCase().contains(keywords)) {
                    return true
                }
                return false
            }
        });
        subFiles.each {
            if (it.isDirectory()) {
                resultList.addAll(searchFileListByLineSizeAndKW(it, keywords))
            } else {
                resultList.add(it)
            }
        }
        return resultList
    }

}

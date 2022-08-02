package com.lvyou.micro.utils.excel;

import cn.hutool.core.collection.CollUtil;
import com.lvyou.micro.exception.ApiException;
import com.lvyou.micro.utils.LocalDateTimeUtils;
import com.lvyou.micro.utils.MathUtils;
import io.swagger.annotations.ApiModelProperty;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellType.BLANK;


/**
 * <p>
 * excel导入工具类
 * </p>
 *
 * @author kun.tan
 * @since 2022-06-06
 */
public class ExcelImportUtil {


    private ExcelImportUtil() {
        //do nothing
    }

    private static final String EXCEL_2003 = ".xls";
    private static final String EXCEL_2007 = ".xlsx";

    /**
     * 解析excel导入数据
     *
     * @param request http请求
     * @return List<List<Object>>   List<Object> 表示行,  object表示列
     * @author kun.tan
     * @date 11:31 2022-07-18
     */
    public static List<List<Object>> getDataListFromExcelFile(HttpServletRequest request) throws FileUploadException,IOException {
        DiskFileItemFactory factory = new DiskFileItemFactory();
        // 设置缓冲区大小，这里是4kb
        factory.setSizeThreshold(4096);
        ServletFileUpload upload = new ServletFileUpload(factory);
        // 解决文件名称乱码
        upload.setHeaderEncoding("utf-8");
        List<FileItem> itemsList = upload.parseRequest(request);

        if (CollectionUtils.isEmpty(itemsList)) {
            throw new ApiException("请选择文件");
        }
        FileItem fileItem = itemsList.get(0);
        // 创建Excel工作薄
        Workbook work = getWorkbook(fileItem);
        // 工作薄转对象集合
        return changeWorkToObjectList(work);
    }

    /**
     * 解析excel导入数据(通过url下载)
     *
     * @param fileName 文件名（带后缀）
     * @param fileUrl 文件地址
     * @return List<List<Object>>   List<Object> 表示行,  object表示列
     * @author kun.tan
     * @date 11:31 2022-07-18
     */
    public static List<List<Object>> getDataListFromExcelFileUrl(String fileName,String fileUrl) throws FileUploadException,IOException {
        Workbook work ;
        //1：验证文件格式
        String fileType = fileName.substring(fileName.lastIndexOf('.'));
        InputStream fileStream= getInputStreamFromUrl(fileUrl);
        if (EXCEL_2003.equals(fileType)) {
            work = new HSSFWorkbook(fileStream);
        } else if (EXCEL_2007.equals(fileType)) {
            work = new XSSFWorkbook(fileStream);
        } else {
            throw new ApiException("解析的文件有误");
        }

        // 工作薄转对象集合
        return changeWorkToObjectList(work);
    }
    /**
     * 获取列名和列的序号
     *
     * @param list 数据
     * @param rowNumOfColumnsName 列名所在的行号(从1开始)
     * @return HashMap<String, Integer>   HashMap<列名, 列序号>
     * @author kun.tan
     * @date 11:31 2022-07-18
     */
    public static HashMap<String, Integer> getColumnNameAndIndex(List<List<Object>> list, Integer rowNumOfColumnsName) {
        HashMap<String,Integer> columnMap=new HashMap<String,Integer>();
        List<Object> rowOfColumnsName=list.get(rowNumOfColumnsName-1);
        Integer columnNum= rowOfColumnsName.size();

        for (int i = 0; i < columnNum; i++) {
            if(StringUtils.isNotEmpty(rowOfColumnsName.get(i).toString())){
                columnMap.put(rowOfColumnsName.get(i).toString().trim(),i);
            }
            continue;
        }
        return columnMap;
    }


    /**
     * 根据文件url导入数据返回对象集合
     *
     * @param clazz 对象类型
     * @param fileName 文件名称
     * @param fileUrl   文件地址
     * @param ignoreStartRowNum   开头忽略的行数
     * @param columnsNameRowNum   列名行号
     * @param dataStartRowNum   数据开始行号
     * @param ignoreEndRowNum   结尾忽略的行数
     * @return List<T>  对象集合
     * @author kun.tan
     * @date 11:31 2022-08-01
     */
    public static <T> List<T> getDataListFromExcelFileUrl(Class<T> clazz,String fileName,String fileUrl,Integer ignoreStartRowNum,
                                                          Integer columnsNameRowNum, Integer dataStartRowNum,Integer ignoreEndRowNum) throws FileUploadException,IOException{
        List<List<Object>> importExcelData= getDataListFromExcelFileUrl(fileName,fileUrl);
        HashMap<String, Integer> columbMap= getColumnNameAndIndex(importExcelData,columnsNameRowNum);
        return changeToObjList(clazz, importExcelData, columbMap,ignoreStartRowNum, dataStartRowNum,ignoreEndRowNum);
    }
    /**
     * 导入文件流返回对象集合
     * @param request 请求体
     * @param clazz 对象类型
     * @param ignoreStartRowNum   开头忽略的行数
     * @param columnsNameRowNum   列名行号
     * @param dataStartRowNum   数据开始行号
     * @param ignoreEndRowNum   结尾忽略的行数
     * @return List<T>  对象集合
     * @author kun.tan
     * @date 11:31 2022-08-01
     */
    public static <T> List<T> getDataListFromExcelFile(HttpServletRequest request,Class<T> clazz,Integer ignoreStartRowNum,
                                                       Integer columnsNameRowNum, Integer dataStartRowNum,Integer ignoreEndRowNum) throws FileUploadException,IOException{
        List<List<Object>> importExcelData= getDataListFromExcelFile(request);
        HashMap<String, Integer> columbMap= getColumnNameAndIndex(importExcelData,columnsNameRowNum);
        return changeToObjList(clazz, importExcelData, columbMap,ignoreStartRowNum, dataStartRowNum,ignoreEndRowNum);
    }

    /**
     * 导入数据转对象集合
     *
     * @param clazz 对象类型
     * @param importExcelData 解析的数据集合
     * @param columbMap   HashMap<列名, 列序号>
     * @param ignoreStartRowNum   开头忽略的行数
     * @param dataStartRowNum   数据开始行号
     * @param ignoreEndRowNum   结尾忽略的行数
     * @return List<T>  对象集合
     * @author kun.tan
     * @date 11:31 2022-08-01
     */
    private static <T> List<T> changeToObjList(Class<T> clazz, List<List<Object>> importExcelData, HashMap<String, Integer> columbMap
            ,Integer ignoreStartRowNum, Integer dataStartRowNum,Integer ignoreEndRowNum) {
        List<T> dtoList = new ArrayList<>();
        if ((long) importExcelData.size() > dataStartRowNum) {
            int i= dataStartRowNum - 1;
            String annotationValue="";
            try {
                for (; i < (long) importExcelData.size() - ignoreEndRowNum; i++) {
                    List<Object> row = importExcelData.get(i);
                    T addDto = clazz.newInstance();
                    for (Field field : getAllFields(clazz)) {
                        field.setAccessible(true);
                        annotationValue= field.getAnnotation(ApiModelProperty.class).value();
                        if(columbMap.get(annotationValue)==null){
                            continue;
                        }
                        if(field.getType().equals(BigDecimal.class)){
                            field.set(addDto,  MathUtils.getBigDecimal(row.get(columbMap.get(annotationValue)).toString()));
                        } else if(field.getType().equals(LocalDateTime.class)){
                            field.set(addDto,  LocalDateTimeUtils.convertTimeStrToLocalDateTime(row.get(columbMap.get(annotationValue)).toString()));
                        }else if(field.getType().equals(LocalDate.class)){
                            field.set(addDto,  LocalDateTimeUtils.convertTimeStrToLocalDate(row.get(columbMap.get(annotationValue)).toString()));
                        } else {
                            field.set(addDto, row.get(columbMap.get(annotationValue)).toString());
                        }

                    }
                    dtoList.add(addDto);
                }
            } catch (Exception e) {
                throw new ApiException("导入数据解析错误：第"+ignoreStartRowNum+i+1+"行（"+annotationValue+" "+e.getMessage()+"）");
            }
        }
        return dtoList;
    }

    private static List<Field> getAllFields(Class clazz) {
        List<Field> fields = new ArrayList<>();
        if (clazz != null) {
            fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
            final List<Field> allFields = getAllFields(clazz.getSuperclass());
            if (CollUtil.isNotEmpty(allFields)) {
                fields.addAll(allFields);
            }
        }
        return fields;
    }


    /**
     * 根据url下载文件流
     * @param urlStr
     * @return
     */
    private static InputStream getInputStreamFromUrl(String urlStr) {
        InputStream inputStream=null;
        try {
            //url解码
            URL url = new URL(java.net.URLDecoder.decode(urlStr, "UTF-8"));
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            //设置超时间为3秒
            conn.setConnectTimeout(3 * 1000);
            //防止屏蔽程序抓取而返回403错误
            conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
            //得到输入流
            inputStream = conn.getInputStream();
        } catch (IOException e) {

        }
        return inputStream;
    }




    private static Workbook getWorkbook(FileItem fileItem) throws IOException {
        Workbook work = null;
        //1：验证文件格式
        String fileName = fileItem.getName();
        String fileType = fileName.substring(fileName.lastIndexOf('.'));
        if (EXCEL_2003.equals(fileType)) {
            work = new HSSFWorkbook(fileItem.getInputStream());
        } else if (EXCEL_2007.equals(fileType)) {
            work = new XSSFWorkbook(fileItem.getInputStream());
        } else {
            throw new ApiException("解析的文件有误");
        }
        return work;
    }

    private static List<List<Object>> changeWorkToObjectList(Workbook work) {
        List<List<Object>> list = new ArrayList<>();
        Sheet sheet ;
        Row row ;
        Cell cell ;
        //遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            int addLineNum = 2;
            //遍历当前sheet中的所有行
            for (int j = sheet.getFirstRowNum()+ addLineNum ; j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (row == null || row.getFirstCellNum() == j || isRowEmpty(row)) {
                    continue;
                }
                //遍历所有的列
                List<Object> li = new ArrayList<>();
                for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
                    cell = row.getCell(k);
                    li.add(cell);
                }
                list.add(li);
            }
        }
        return list;
    }

    private static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != BLANK) {
                return false;
            }
        }
        return true;
    }

}

package util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

public class ExcelUtils {

    /**
     * 简单读，默认行头为1
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param readListener 监听器
     */
    public static void simpleRead(String path, Class clazz, ReadListener readListener) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        EasyExcel.read(path, clazz, readListener).sheet().doRead();
    }

    public static void simpleRead(InputStream is, Class clazz, ReadListener readListener) {
        if(is==null || readListener==null){
            return;
        }
        EasyExcel.read(is, clazz, readListener).sheet().doRead();
    }

    /**
     * 读取多行头Excel
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param readListener 监听器
     * @param rownum 行头数
     */
    public static void complexHeaderRead(String path, Class clazz, ReadListener readListener, int rownum) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        if(rownum<1){
            rownum=0;
        }
        EasyExcel.read(path, clazz, readListener).sheet().headRowNumber(rownum).doRead();
    }

    public static void complexHeaderRead(InputStream is, Class clazz, ReadListener readListener, int rownum) {
        if(is==null || readListener==null){
            return;
        }
        if(rownum<1){
            rownum=0;
        }
        EasyExcel.read(is, clazz, readListener).sheet().headRowNumber(rownum).doRead();
    }

    /**
     * 根据下标和列名读取Excel
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param readListener 监听器
     */
    public static void indexOrNameRead(String path, Class clazz, ReadListener readListener) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        EasyExcel.read(path, clazz, readListener).sheet().doRead();
    }

    public static void indexOrNameRead(InputStream is, Class clazz, ReadListener readListener) {
        if(is==null || readListener==null){
            return;
        }
        EasyExcel.read(is, clazz, readListener).sheet().doRead();
    }

    /**
     * 读取所有Sheet页数据
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param readListener 监听器
     */
    public static void repeatedAllRead(String path, Class clazz, ReadListener readListener) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        EasyExcel.read(path, clazz, readListener).doReadAll();
    }

    public static void repeatedAllRead(InputStream is, Class clazz, ReadListener readListener) {
        if(is==null || readListener==null){
            return;
        }
        EasyExcel.read(is, clazz, readListener).doReadAll();
    }

    /**
     * 读取指定Sheet页数据，注意实体映射类、监听器、Sheet下标必须对应
     * @param path 文件路径
     * @param clazzs 映射类数组
     * @param readListeners 监听器数组
     * @param sheets Sheet页下标数组
     */
    public static void repeatedRead(String path, Class[] clazzs, ReadListener[] readListeners, int[] sheets) {
        if(!new File(path).exists() || clazzs==null || readListeners==null || sheets==null || clazzs.length<1 || readListeners.length<1 || sheets.length<1){
            return;
        }
        if(clazzs.length!=readListeners.length || clazzs.length!=sheets.length){
            return;
        }
        ExcelReader excelReader = EasyExcel.read(path).build();
        List<ReadSheet> readSheets = new ArrayList<ReadSheet>();
        for(int i=0,j=clazzs.length; i<j; i++){
            readSheets.add(EasyExcel.readSheet(sheets[i]).head(clazzs[i]).registerReadListener(readListeners[i]).build());
        }
        excelReader.read(readSheets);
        //必须关闭，读的时候会创建临时文件，不关闭临时文件会占满磁盘
        excelReader.finish();
    }

    public static void repeatedRead(InputStream is, Class[] clazzs, ReadListener[] readListeners, int[] sheets) {
        if(is==null || clazzs==null || readListeners==null || sheets==null || clazzs.length<1 || readListeners.length<1 || sheets.length<1){
            return;
        }
        if(clazzs.length!=readListeners.length || clazzs.length!=sheets.length){
            return;
        }
        ExcelReader excelReader = EasyExcel.read(is).build();
        List<ReadSheet> readSheets = new ArrayList<ReadSheet>();
        for(int i=0,j=clazzs.length; i<j; i++){
            readSheets.add(EasyExcel.readSheet(sheets[i]).head(clazzs[i]).registerReadListener(readListeners[i]).build());
        }
        excelReader.read(readSheets);
        //必须关闭，读的时候会创建临时文件，不关闭临时文件会占满磁盘
        excelReader.finish();
    }

    /**
     * 不创建映射对象的读，不是特别推荐使用，都用String接收对日期的支持不是很好
     * @param path 文件路径
     * @param readListener 监听器
     */
    public static void noModleRead(String path, ReadListener readListener) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        EasyExcel.read(path, readListener).sheet().doRead();
    }

    public static void noModleRead(InputStream is, ReadListener readListener) {
        if(is==null || readListener==null){
            return;
        }
        EasyExcel.read(is, readListener).sheet().doRead();
    }

    /**
     * 读取列名
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param readListener 监听器
     */
    public static void headerRead(String path, Class clazz, ReadListener readListener) {
        if(!new File(path).exists() || readListener==null){
            return;
        }
        EasyExcel.read(path, clazz, readListener).sheet().doRead();
    }

    public static void headerRead(InputStream is, Class clazz, ReadListener readListener) {
        if(is==null || readListener==null){
            return;
        }
        EasyExcel.read(is, clazz, readListener).sheet().doRead();
    }

    /**
     * 简单写
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     */
    public static void simpleWrite(String path, Class clazz, List datas) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).sheet().doWrite(datas);
    }

    public static void simpleWrite(OutputStream os, Class clazz, List<?> datas) {
        if(os==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).sheet().doWrite(datas);
    }

    /**
     * 写入时忽略指定字段
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     * @param excludeColumnFiledNames 指定的忽略字段
     */
    public static void excludeWrite(String path, Class clazz, List<?> datas, Set<String> excludeColumnFiledNames) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).excludeColumnFiledNames(excludeColumnFiledNames).sheet().doWrite(datas);
    }

    public static void excludeWrite(OutputStream os, Class clazz, List<?> datas, Set<String> excludeColumnFiledNames) {
        if(os==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).excludeColumnFiledNames(excludeColumnFiledNames).sheet().doWrite(datas);
    }

    /**
     * 只写入指定的字段
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     * @param includeColumnFiledNames 写入的指定字段
     */
    public static void includeWrite(String path, Class clazz, List<?> datas, Set<String> includeColumnFiledNames) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).includeColumnFiledNames(includeColumnFiledNames).sheet().doWrite(datas);
    }

    public static void includeWrite(OutputStream os, Class clazz, List<?> datas, Set<String> includeColumnFiledNames) {
        if(os==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).includeColumnFiledNames(includeColumnFiledNames).sheet().doWrite(datas);
    }

    /**
     * 写入指定列
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     */
    public static void indexWrite(String path, Class clazz, List<?> datas) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).sheet().doWrite(datas);
    }

    public static void indexWrite(OutputStream os, Class clazz, List<?> datas) {
        if(os==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).sheet().doWrite(datas);
    }

    /**
     * 写入多行头Excel
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     */
    public static void complexHeadWrite(String path, Class clazz, List<?> datas) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).sheet().doWrite(datas);
    }

    public static void complexHeadWrite(OutputStream os, Class clazz, List<?> datas) {
        if(os==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).sheet().doWrite(datas);
    }

    /**
     * 写入多Sheet页Excel
     * @param path 文件路径
     * @param clazzs 实体映射类
     * @param dataGroups 数据
     */
    public static void repeatedWrite(String path, List<Class> clazzs, List<List<?>> dataGroups) {
        if(StringUtils.isEmpty(path) || clazzs==null || clazzs.size()<1 || dataGroups==null || dataGroups.size()<1){
            return;
        }
        if(clazzs.size()!=dataGroups.size()){
            return;
        }
        ExcelWriter excelWriter = EasyExcel.write(path).build();
        for (int i=0,j=clazzs.size(); i<j; i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(i).head(clazzs.get(i)).build();
            excelWriter.write(dataGroups.get(i), writeSheet);
        }
        //关闭流
        excelWriter.finish();
    }

    public static void repeatedWrite(OutputStream os, List<Class> clazzs, List<List<?>> dataGroups) {
        if(os==null || clazzs==null || clazzs.size()<1 || dataGroups==null || dataGroups.size()<1){
            return;
        }
        if(clazzs.size()!=dataGroups.size()){
            return;
        }
        ExcelWriter excelWriter = EasyExcel.write(os).build();
        for (int i=0,j=clazzs.size(); i<j; i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(i).head(clazzs.get(i)).build();
            excelWriter.write(dataGroups.get(i), writeSheet);
        }
        //关闭流
        excelWriter.finish();
    }

    /**
     * 根据模板写入Excel
     * @param path 文件路径
     * @param templatePath 模板路径
     * @param clazz 实体映射类
     * @param datas 数据
     */
    public static void templateWrite(String path, String templatePath, Class clazz, List<?> datas) {
        if(StringUtils.isEmpty(path) || StringUtils.isEmpty(templatePath) || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path, clazz).withTemplate(templatePath).sheet().doWrite(datas);
    }

    public static void templateWrite(OutputStream os, InputStream is, Class clazz, List<?> datas) {
        if(os==null || is==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os, clazz).withTemplate(is).sheet().doWrite(datas);
    }

    /**
     * 自定义样式Excel写入
     * @param path 文件路径
     * @param clazz 实体映射类
     * @param datas 数据
     * @param horizontalCellStyleStrategy 样式策略
     */
    public static void styleWrite(String path, Class clazz, List<?> datas, HorizontalCellStyleStrategy horizontalCellStyleStrategy) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null ){
            return;
        }
        EasyExcel.write(path, clazz).registerWriteHandler(horizontalCellStyleStrategy).sheet().doWrite(datas);
    }

    public static void styleWrite(OutputStream os, Class clazz, List<?> datas, HorizontalCellStyleStrategy horizontalCellStyleStrategy) {
        if(os==null || clazz==null || datas==null ){
            return;
        }
        EasyExcel.write(os, clazz).registerWriteHandler(horizontalCellStyleStrategy).sheet().doWrite(datas);
    }

    /**
     * 单元格合并Excel写入
     * @param path 文件路径
     * @param clazz 实体映射策略
     * @param datas 数据
     * @param loopMergeStrategy 合并策略
     */
    public static  void mergeWrite(String path, Class clazz, List<?> datas, LoopMergeStrategy loopMergeStrategy) {
        if(StringUtils.isEmpty(path) || clazz==null || datas==null ){
            return;
        }
        EasyExcel.write(path, clazz).registerWriteHandler(loopMergeStrategy).sheet().doWrite(datas);
    }

    public static void mergeWrite(OutputStream os, Class clazz, List<?> datas, LoopMergeStrategy loopMergeStrategy) {
        if(os==null || clazz==null || datas==null ){
            return;
        }
        EasyExcel.write(os, clazz).registerWriteHandler(loopMergeStrategy).sheet().doWrite(datas);
    }

    /**
     * 动态行头写入Excel
     * @param path 文件路径
     * @param heads 动态行头
     * @param clazz 实体映射类
     * @param datas 数据
     */
    public static void dynamicHeadWrite(String path, List<List<String>> heads, Class clazz, List<?> datas) {
        if(StringUtils.isEmpty(path) || heads==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(path).head(heads).sheet().doWrite(datas);
    }

    public static void dynamicHeadWrite(OutputStream os, List<List<String>> heads, Class clazz, List<?> datas) {
        if(os==null || heads==null || clazz==null || datas==null){
            return;
        }
        EasyExcel.write(os).head(heads).sheet().doWrite(datas);
    }

    /**
     * 无实体映射类写入
     * @param path 文件路径
     * @param heads 行头
     * @param datas 数据
     */
    public static void noModleWrite(String path, List<List<String>> heads, List<List<?>> datas) {
        if(StringUtils.isEmpty(path) || heads==null || datas==null){
            return;
        }
        EasyExcel.write(path).head(heads).sheet().doWrite(datas);
    }

    public static void noModleWrite(OutputStream os, List<List<String>> heads, List<List<?>> datas) {
        if(os==null || heads==null || datas==null){
            return;
        }
        EasyExcel.write(os).head(heads).sheet().doWrite(datas);
    }

    /**
     * description 动态下载模板
     * 下载模板时，动态获取表头，无data数据
     * @param path
     * @param heads
     * @return
     * @throws
     * @author lizl
     * @date 2020/3/5 18:41
     */
    public static void downloadTemplate(String path, List<List<String>> heads){
        if(StringUtils.isBlank(path) || heads==null){
            return;
        }
        EasyExcel.write(path).sheet().doWrite(heads);
    }

    /**
     * description 动态下载模板
     * 下载模板时，动态获取表头，无data数据
     * @param os
     * @param heads
     * @return
     * @throws
     * @author lizl
     * @date 2020/3/5 18:41
     */
    public static void downloadTemplate(OutputStream os, List<List<String>> heads){
        if(os==null || heads==null){
            return;
        }
        EasyExcel.write(os).sheet().doWrite(heads);
    }
}

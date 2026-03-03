package org.murchey.newhandlethebill.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.Setter;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;

@Controller
@RequestMapping("/api")
@CrossOrigin(origins = "*")//允许跨域
public class ExcelUtils {
    public static Map<String, Double> ReadExcel(File excel,
                                                int headRow,
                                                int nameIndex,
                                                int moneyIndex) {
        Map<String, Double> result = new LinkedHashMap<>();

        EasyExcel.read(excel, new AnalysisEventListener<Map<Integer,String>>() {

            private boolean headSkipped = false;
            //AnalysisContext 上下文分析器
            @Override
            public void invoke(Map<Integer, String> row, AnalysisContext ctx) {
                if (!headSkipped && ctx.readRowHolder().getRowIndex() < headRow) {
                    return; // 跳过表头行及之前的所有行
                } else if (!headSkipped) {
                    headSkipped = true; // 表头行已处理完毕
                }

                String name = row.getOrDefault(nameIndex, "").trim();
                if (name.isEmpty()) return;

                double money = 0;
                try {
                    money = Double.parseDouble(row.getOrDefault(moneyIndex, "0").replaceAll(",", ""));
                } catch (NumberFormatException ignore) {}

                result.merge(name, money, Double::sum);
            }

            @Override public void onException(Exception e, AnalysisContext ctx) {}
            @Override public void doAfterAllAnalysed(AnalysisContext ctx) {}
        }).sheet().doRead();
        return result;
    }

    //名单差异比对方法
    public static List<List<String>> FindNameDiff(Map<String, Double> handleExcel, Map<String, Double> standardExcel){
        Set<String> handleName = handleExcel.keySet();
        Set<String> standardName = standardExcel.keySet();
        //原表格两个的副本
        Set<String> namesOnlyInHandleExcel = new HashSet<>(handleName);
        Set<String> namesOnlyInStandardExcel = new HashSet<>(standardName);

        //handle - standard是只在待处理表格里有的 -> 错付或多付款名单（标准账单中没有此人）
        namesOnlyInHandleExcel.removeAll(standardName);
        //standard - handle是只在标准表格里有的 -> 未付款名单（付款账单中没有此人）
        namesOnlyInStandardExcel.removeAll(handleName);

        //生成写入行表格
        List<String> namesOnlyInHandleLs = new ArrayList<>(namesOnlyInHandleExcel);
        List<String> namesOnlyInStandardLs = new ArrayList<>(namesOnlyInStandardExcel);
        int max=Math.max(
                namesOnlyInHandleLs.size(),namesOnlyInStandardLs.size()
        );
        List<List<String>> rows = new ArrayList<>(max+1);
        rows.add(Arrays.asList("错付或多付款名单（标准账单中没有此人）","未付款名单（付款账单中没有此人）"));
        for (int i = 0; i < max; i++) {
            String grid1 = i < namesOnlyInHandleLs.size()?namesOnlyInHandleLs.get(i):"";
            String grid2 = i < namesOnlyInStandardLs.size()?namesOnlyInStandardLs.get(i):"";
            rows.add(Arrays.asList(grid1,grid2));
        }
        return rows;
    }
    //单元格差异对比
    public static final BigDecimal DEFAULT_TOL = new BigDecimal("0.01");
    public static final class DiffResult{
        public final Set<String> onlyHandle;
        public final Set<String> onlyStandard;
        public final Map<String,MoneyDiff> diffMoney;
        private DiffResult(
                Set<String> onlyHandle,
                Set<String> onlyStandard,
                Map<String,MoneyDiff> diffMoney
        ){
            this.onlyHandle=Collections.unmodifiableSet(onlyHandle);
            this.onlyStandard=Collections.unmodifiableSet(onlyStandard);
            this.diffMoney=Collections.unmodifiableMap(diffMoney);
        }
    }
    public static final class MoneyDiff{
        public final BigDecimal handle;
        public final BigDecimal standard;
        public final BigDecimal diff;
        MoneyDiff(BigDecimal h,BigDecimal s){
            this.standard=s;
            this.handle=h;
            this.diff = s.subtract(h);
        }
    }
    private static BigDecimal toBD(Double d) {
        return d == null ? BigDecimal.ZERO : BigDecimal.valueOf(d);
    }
    public static DiffResult FindGridDiff(Map<String, Double> handleExcel, Map<String, Double> standardExcel, BigDecimal tolerance){
        Set<String> onlyHandle = new HashSet<>(handleExcel.keySet());
        Set<String> onlyStandard = new HashSet<>(standardExcel.keySet());
        Set<String> common = new HashSet<>(handleExcel.keySet());
        Map<String,MoneyDiff> diffMoney = new LinkedHashMap<>();
        // 计算交集：同时存在于两个Excel中的键
        common.retainAll(standardExcel.keySet());
        for (String item:common){
            BigDecimal h = toBD(handleExcel.get(item));
            BigDecimal s = toBD(standardExcel.get(item));
            if (h.subtract(s).abs().compareTo(tolerance)>0){
                diffMoney.put(item,new MoneyDiff(h,s));
            }
        }
        return new DiffResult(onlyHandle,onlyStandard,diffMoney);
    }
    public static int LetterToIndex(String str){
        char ch = str.charAt(0);
        return Character.toUpperCase(ch)-'A';//直接转换成excel的列索引
    }
    @Setter
    private static int Progress = 0;

    @GetMapping("/getProgress")
    @ResponseBody
    public double getProgress(){
        return Progress/100.0;
    }

    //对比主方法
    @GetMapping("/startCmp")
    @ResponseBody
    public static boolean BatchCompareAndSave(@RequestParam String needHandleFilePath,
                                              @RequestParam String standardFilePath,
                                              @RequestParam int handleHeadRow,
                                              @RequestParam String handleNameCol,
                                              @RequestParam String handleMoneyCol,
                                              @RequestParam int standardHeadRow,
                                              @RequestParam String standardNameCol,
                                              @RequestParam String standardMoneyCol,
                                              @RequestParam String resultPath) {

        List<String[]> pairs = GetPathUtils.findSameNamePairs(
                new File(needHandleFilePath),
                new File(standardFilePath));

        if (pairs.isEmpty()) return false;

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        File totalFile = new File(resultPath, "总对比结果_" + timestamp + ".xlsx");

        /* 一个 writer 贯穿全部比对 */
        try (ExcelWriter writer = EasyExcel.write(totalFile).build()) {

            int total = pairs.size();

            for (int i = 0; i < total; i++) {
                String[] p = pairs.get(i);          // 1. 取当前这一对
                if (p == null || p.length < 3) continue; // 跳过无效的数组
                File fileA = new File(p[1]);
                File fileB = new File(p[2]);
                String prefix = p[0];                 // 2. 原始文件名（当 Sheet 名）

                /* 读数据 */
                Map<String, Double> handleData =
                        ReadExcel(fileA,
                                handleHeadRow,
                                LetterToIndex(handleNameCol),
                                LetterToIndex(handleMoneyCol));

                Map<String, Double> standardData =
                        ReadExcel(fileB,
                                standardHeadRow,
                                LetterToIndex(standardNameCol),
                                LetterToIndex(standardMoneyCol));

                /* 计算差异 */
                List<List<String>> nameRows = FindNameDiff(handleData, standardData);
                DiffResult grid = FindGridDiff(handleData, standardData,
                        DEFAULT_TOL);

                /* 拼成同一张表 */
                List<List<String>> oneSheet = new ArrayList<>();
                oneSheet.add(Arrays.asList("==== 名单差异 ===="));
                oneSheet.addAll(nameRows);

                oneSheet.add(Arrays.asList("", ""));   // 空行分隔

                oneSheet.add(Arrays.asList("==== 金额差异 ===="));
                oneSheet.add(Arrays.asList("姓名", "待处理金额", "标准金额", "差值"));
                grid.diffMoney.forEach((k, v) ->
                        oneSheet.add(Arrays.asList(k,
                                v.handle.toPlainString(),
                                v.standard.toPlainString(),
                                v.diff.toPlainString())));

                /* 写进以文件名命名的 Sheet */
                WriteSheet sheet = EasyExcel.writerSheet(prefix).build();
                writer.write(oneSheet, sheet);

                setProgress((int) ((i + 1) * 100.0 / total));
            }
        }
        return true;
    }
}
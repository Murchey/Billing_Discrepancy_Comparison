package org.example.ManyUtils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import javax.sound.sampled.Line;
import javax.swing.*;
import java.awt.Component;
import java.awt.Frame;
import java.io.File;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

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
    //名单差异储存方法
//    private static void SaveNameDiffsExcel(String resExcelPath, String resExcelSheet,List<List<String>> rows){
//        EasyExcel.write(resExcelPath)
//                .sheet(resExcelSheet)
//                .doWrite(rows);
//    }
    //单元格差异比对方法
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
//    private static void SaveGridDiffsExcel(String filePath,
//                                           String sheetName,
//                                           DiffResult res) {
//
//        List<List<Object>> rows = new ArrayList<>();
//        rows.add(Arrays.asList("姓名", "待处理金额", "标准金额", "差值"));
//
//        res.diffMoney.forEach((k, v) -> rows.add(
//                Arrays.asList(k,
//                        v.handle.doubleValue(),
//                        v.standard.doubleValue(),
//                        v.diff.doubleValue())));
//
//        EasyExcel.write(filePath)          // 新建或覆盖
//                .sheet(sheetName)         // 第二个 sheet
//                .doWrite(rows);
//    }
    public static int LetterToIndex(String str){
        char ch = str.charAt(0);
        return Character.toUpperCase(ch)-'A';//直接转换成excel的列索引
    }
    //外部可调用方法

    /**
     * 批量比对同名文件对，并生成总对比结果文件
     * @param needHandleFilePath 待核对账单文件夹路径
     * @param standardFilePath 标准账单文件夹路径
     * @param handleHeadRow 待核对账单表头行
     * @param handleNameCol 待核对账单姓名列
     * @param handleMoneyCol 待核对账单金额列
     * @param standardHeadRow 标准账单表头行
     * @param standardNameCol 标准账单姓名列
     * @param standardMoneyCol 标准账单金额列
     * @param parentComponent 父组件，用于显示对话框
     * @return true-成功；false-出错
     */
    public static boolean BatchCompareAndSave(String needHandleFilePath,
                                             String standardFilePath,
                                             int handleHeadRow,
                                             String handleNameCol,
                                             String handleMoneyCol,
                                             int standardHeadRow,
                                             String standardNameCol,
                                             String standardMoneyCol,
                                             Component parentComponent) {
        /* 1. 进度窗 */
        JDialog progressDlg = new JDialog((Frame) SwingUtilities.getWindowAncestor(parentComponent),
                "正在比对...", true);
        JProgressBar bar = new JProgressBar(0, 100);
        bar.setStringPainted(true);
        progressDlg.add(bar);
        progressDlg.setSize(300, 80);
        progressDlg.setLocationRelativeTo(parentComponent);

        /* 2. 后台任务 */
        SwingWorker<Boolean, Integer> worker = new SwingWorker<>() {
            @Override
            protected Boolean doInBackground() throws Exception {
                List<String[]> pairs = GetPathUtils.findSameNamePairs(
                        new File(needHandleFilePath),
                        new File(standardFilePath));

                if (pairs.isEmpty()) return false;

                File resultDir = GetPathUtils.buildResultFile().getParentFile();
                if (!resultDir.exists()) resultDir.mkdirs();
                String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
                File totalFile = new File(resultDir, "总对比结果_" + timestamp + ".xlsx");

                /* 一个 writer 贯穿全部比对 */
                try (ExcelWriter writer = EasyExcel.write(totalFile).build()) {

                    int total = pairs.size();
                    setProgress(0);

                    for (int i = 0; i < total; i++) {
                        String[] p   = pairs.get(i);          // 1. 取当前这一对
                        if (p == null || p.length < 3) continue; // 跳过无效的数组
                        File fileA   = new File(p[1]);
                        File fileB   = new File(p[2]);
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
                        DiffResult grid  = FindGridDiff(handleData, standardData,
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

            @Override
            protected void done() {
                progressDlg.dispose();
                try {
                    boolean success = get();          // 捕获异常
                    if (success) {
                        JOptionPane.showMessageDialog(parentComponent,
                                "全部比对完成！\n请将目录的结果文件复制保存，软件退出后会自动删除。", "完成", JOptionPane.INFORMATION_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(parentComponent,
                                "未找到同名文件对！", "提示", JOptionPane.INFORMATION_MESSAGE);
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(parentComponent,
                            "比对出错：" + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        };

        /* 3. 把进度实时刷到进度条 */
        worker.addPropertyChangeListener(evt -> {
            if ("progress".equals(evt.getPropertyName())) {
                bar.setValue((Integer) evt.getNewValue());
            }
        });

        /* 4. 启动 */
        worker.execute();
        progressDlg.setVisible(true);   // 模态阻塞，直到 worker.done() 关闭
        return true;
    }
}

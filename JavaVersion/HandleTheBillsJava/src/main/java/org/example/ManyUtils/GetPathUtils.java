package org.example.ManyUtils;

import org.example.Main;

import javax.swing.*;
import java.io.File;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class GetPathUtils {
    private static final List<String> directoryNames = Arrays.asList("tempData","result");
    public static void PathInit(){
        File appDir = new File(
                Main.class.getProtectionDomain()
                        .getCodeSource()
                        .getLocation()
                        .getPath())
                .getParentFile();
        directoryNames.forEach(name -> {
            File dir = new File(appDir, name);
            if (!dir.exists()) {          // 不存在就建
                boolean ok = dir.mkdir(); // 只建一层
                System.out.println(ok ? "已创建：" + name : "创建失败：" + name);
            }
        });
    }
    public static File GetMyPath(){
        try {
            // 获取当前类的位置
            String path = Main.class.getProtectionDomain().getCodeSource().getLocation().getPath();
            File location = new File(path);
            
            // 检查是否是从jar文件中运行
            if (location.getName().endsWith(".jar")) {
                // 如果是jar文件，返回jar文件所在的目录
                return location.getParentFile();
            } else if (location.isDirectory()) {
                // 如果是目录（开发环境），返回当前目录
                return location;
            }
        } catch (Exception e) {
            // 如果获取失败，使用用户当前工作目录
            e.printStackTrace();
        }
        
        // 默认使用用户当前工作目录
        return new File(System.getProperty("user.dir"));
    }

    public static File buildResultFile() {
        // 1. 软件根目录
        File root = GetMyPath();
        // 2. result 子目录
        File resultDir = new File(root, "result");
        if (!resultDir.exists()) resultDir.mkdirs();          // 自动创建
        // 3. 时间戳文件名
        String time = LocalDateTime.now()
                .format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String fileName = time + "_比对结果.xlsx";
        return new File(resultDir, fileName);
    }

    public static File SelectDirPath(){
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int res = fc.showOpenDialog(null);
        if (res == JFileChooser.APPROVE_OPTION){
            return fc.getSelectedFile();
        }else{
            System.out.println("选择取消，无路径返回");
            return new File("no path");
        }
    }
    //此方法遍历目录
    public static String[] scanFiles(String dirPath, String fileExtention){
        File directory = new File(dirPath);
        if (!directory.exists()||!directory.isDirectory()){
            throw new IllegalArgumentException("目录不存在或不是有效目录："+dirPath);
        }
        List<String> filePaths = new ArrayList<>();
        scanDirectory(directory,fileExtention,filePaths);
        return filePaths.toArray(new String[0]);
    }

    private static void scanDirectory(File directory,String targetExtension,List<String> filePaths){
        //遍历目录核心方法
        File[] files = directory.listFiles();
        if (files == null){
            return;
        }
        for (File file : files){
            if (file.isDirectory()){
                scanDirectory(file,targetExtension,filePaths);
            }else{
                String fileName = file.getName().toLowerCase(Locale.ROOT);
                if (fileName.endsWith("."+targetExtension.toLowerCase(Locale.ROOT))){
                    filePaths.add(file.getAbsolutePath());
                }
            }
        }
    }
    public static File buildResultDir() {
        File dir = new File(System.getProperty("user.dir"), "对比结果");
        dir.mkdirs();          // 目录不存在就建
        return dir;
    }
    public static List<String[]> findSameNamePairs(File dirA, File dirB) {
        List<String[]> pairs = new ArrayList<>();

        List<File> listA = Arrays.asList(Objects.requireNonNull(dirA.listFiles()));
        List<File> listB = Arrays.asList(Objects.requireNonNull(dirB.listFiles()));

        Map<String, List<File>> mapB = listB.stream()
                .collect(Collectors.groupingBy(File::getName));

        for (File a : listA) {
            List<File> sameInB = mapB.get(a.getName());
            if (sameInB != null) {
                for (File b : sameInB) {
                    pairs.add(new String[]{a.getName(), a.getPath(), b.getPath()});
                }
            }
        }
        return pairs;
    }

}

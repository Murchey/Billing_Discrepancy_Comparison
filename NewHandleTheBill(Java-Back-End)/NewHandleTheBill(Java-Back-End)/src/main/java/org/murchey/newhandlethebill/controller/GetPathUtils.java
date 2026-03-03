package org.murchey.newhandlethebill.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

@Controller
@RequestMapping("/api")
@CrossOrigin(origins = "*")//允许跨域
public class GetPathUtils {
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
    @GetMapping("/openResultDir")
    @ResponseBody
    public void OpenResultDir(@RequestParam String path){
        try {
            String os = System.getProperty("os.name").toLowerCase();
            ProcessBuilder pb;

            if (os.contains("win")) {
                // Windows
                pb = new ProcessBuilder("explorer.exe", path);
            } else if (os.contains("mac")) {
                // macOS
                pb = new ProcessBuilder("open", path);
            } else {
                // Linux
                pb = new ProcessBuilder("xdg-open", path);
            }

            pb.start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

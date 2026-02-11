package org.example.ManyUtils;

import com.formdev.flatlaf.FlatDarkLaf;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import java.awt.*;
import java.io.File;
import java.util.LinkedHashMap;
import java.util.Map;

public class FrameUtils {
    public static Dimension GetScreenSize(JFrame frame){
        //获得屏幕尺寸的方法
        GraphicsConfiguration gc = frame.getGraphicsConfiguration();
        Rectangle bounds = gc.getBounds();
        return new Dimension(bounds.width, bounds.height);
    }

    public static Dimension AdaptFrameSize(Dimension bounds, double heightRate, double widthRate){
        //获得窗口占屏幕总尺寸的方法
        return new Dimension((int)(widthRate*bounds.getWidth()),(int)(heightRate*bounds.getHeight()));
    }

    public static JFrame CreateFrame(String title){
        //创建风格统一、能关闭的窗口的方法
        FlatDarkLaf.setup();
        JFrame f=new JFrame(title);
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);//关闭按钮生效
        return f;
    }

    public static Map<String, JPanel> CreateTabbedPane(JFrame frame, String[] paneNames){
        JTabbedPane pane = new JTabbedPane();
        Map<String, JPanel> pages = new LinkedHashMap<>();
        for (String item : paneNames){
            JPanel tab = new JPanel(new BorderLayout());
            pane.addTab(item,tab);
            pages.put(item,tab);
        }
        frame.add(pane);
        return pages;
    }

    public static void UpdateListModel(JList<File> files,
                                DefaultListModel<File> listModel,
                                String[] paths
    ){
        listModel.clear();
        for (String path : paths){
            File file = new File(path);
            if (file.exists()){
                listModel.addElement(file);
            }else{
                System.err.println("路径不存在："+path);
            }
        }
    }

    public static void OnTextChange(JTextField tf, Runnable callback){
        tf.getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                callback.run();
            }

            @Override
            public void removeUpdate(DocumentEvent e) {
                callback.run();
            }

            @Override
            public void changedUpdate(DocumentEvent e) {
                callback.run();
            }
        });
    }

}

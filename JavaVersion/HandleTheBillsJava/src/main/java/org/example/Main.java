package org.example;

import org.example.ManyUtils.ExcelUtils;
import org.example.ManyUtils.FrameUtils;
import org.example.ManyUtils.GetPathUtils;
import org.example.ManyUtils.VarUtils;

import javax.swing.*;
import javax.swing.plaf.FontUIResource;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.util.Map;

public class Main {
    static FontUIResource res = new FontUIResource("Microsoft YaHei", Font.PLAIN, 19);
    public static void main(String[] args) {
        GetPathUtils.PathInit();
        for (Object key : UIManager.getLookAndFeelDefaults().keySet()) {
            if (key.toString().endsWith(".font")) {
                UIManager.put(key, res);
            }
        }
        JFrame mainFrame = FrameUtils.CreateFrame("账单处理系统-Java重制版");
        mainFrame.setSize(
                FrameUtils.AdaptFrameSize(FrameUtils.GetScreenSize(mainFrame), 0.5, 0.6)
        );
        mainFrame.setLocationRelativeTo(null);//居中
        String[] paneNamesOfMainFrame = {"待核对账单设置", "标准账单设置", "数据比对操作面板"};
        Map<String, JPanel> mainTabs = FrameUtils.CreateTabbedPane(mainFrame, paneNamesOfMainFrame);
    //待核对账单设置页面
    {
        mainTabs.get("待核对账单设置").setLayout(new BoxLayout(
                mainTabs.get("待核对账单设置"), BoxLayout.Y_AXIS
        ));
        //上下边距
        mainTabs.get("待核对账单设置").add(Box.createVerticalStrut(80));

        JPanel pathSettingPanel = new JPanel(new FlowLayout());
        mainTabs.get("待核对账单设置").add(pathSettingPanel);
        JTextField pathText = new JTextField();
        pathText.setPreferredSize(new Dimension(300, 30));
        pathSettingPanel.add(pathText);
        JButton chooseLocationBtn = new JButton("选取待核对账单所在位置");
        chooseLocationBtn.setPreferredSize(new Dimension(250, 30));
        //待核对账单设置页面按钮事件
        {
            chooseLocationBtn.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    pathText.setText(GetPathUtils.SelectDirPath().toString());
                    VarUtils.needHandleFilePath = pathText.getText();
                    System.out.println(VarUtils.needHandleFilePath);
                }
            });
        }
        pathSettingPanel.add(chooseLocationBtn);

        JPanel headRowSettingPanel = new JPanel(new FlowLayout());
        JTextField headRowText = new JTextField();
        headRowText.setPreferredSize(new Dimension(100, 30));
        headRowSettingPanel.add(new JLabel("表头所在行："));
        headRowSettingPanel.add(headRowText);
        FrameUtils.OnTextChange(headRowText,()->{
            VarUtils.handleHeadRow =  Integer.parseInt(headRowText.getText());
            System.out.println("headRowText："+ VarUtils.handleHeadRow);
        });
        mainTabs.get("待核对账单设置").add(headRowSettingPanel);

        JPanel nameColSettingPanel = new JPanel(new FlowLayout());
        JTextField nameColText = new JTextField();
        nameColText.setPreferredSize(new Dimension(100, 30));
        nameColSettingPanel.add(new JLabel("姓名所在列："));
        nameColSettingPanel.add(nameColText);
        FrameUtils.OnTextChange(nameColText,()->{
            VarUtils.handleNameCol = nameColText.getText();
            System.out.println("nameColText："+ VarUtils.handleNameCol);
        });
        mainTabs.get("待核对账单设置").add(nameColSettingPanel);

        JPanel moneyColSettingPanel = new JPanel(new FlowLayout());
        JTextField moneyColText = new JTextField();
        moneyColText.setPreferredSize(new Dimension(100, 30));
        moneyColSettingPanel.add(new JLabel("金额所在列："));
        moneyColSettingPanel.add(moneyColText);
        FrameUtils.OnTextChange(moneyColText,()->{
            VarUtils.handleMoneyCol = moneyColText.getText();
            System.out.println("moneyColText："+ VarUtils.handleMoneyCol);
        });
        mainTabs.get("待核对账单设置").add(moneyColSettingPanel);
    }
    //标准账单设置页面
    {
        mainTabs.get("标准账单设置").setLayout(new BoxLayout(
                mainTabs.get("标准账单设置"), BoxLayout.Y_AXIS
        ));
        //上下边距
        mainTabs.get("标准账单设置").add(Box.createVerticalStrut(80));

        JPanel standardPathSettingPanel = new JPanel(new FlowLayout());
        mainTabs.get("标准账单设置").add(standardPathSettingPanel);
        JTextField standardPathText = new JTextField();
        standardPathText.setPreferredSize(new Dimension(300, 30));
        standardPathSettingPanel.add(standardPathText);
        JButton standardChooseLocationBtn = new JButton("选取标准账单所在位置");
        standardChooseLocationBtn.setPreferredSize(new Dimension(250, 30));
        //标准账单设置页面按钮事件
        {
            standardChooseLocationBtn.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    standardPathText.setText(GetPathUtils.SelectDirPath().toString());
                    VarUtils.standardFilePath=standardPathText.getText();
                    System.out.println(VarUtils.standardFilePath);
                }
            });
        }
        standardPathSettingPanel.add(standardChooseLocationBtn);

        JPanel standardHeadRowSettingPanel = new JPanel(new FlowLayout());
        JTextField standardHeadRowText = new JTextField();
        standardHeadRowText.setPreferredSize(new Dimension(100, 30));
        standardHeadRowSettingPanel.add(new JLabel("表头所在行："));
        standardHeadRowSettingPanel.add(standardHeadRowText);
        FrameUtils.OnTextChange(standardHeadRowText,()->{
            VarUtils.standardHeadRow = Integer.parseInt(standardHeadRowText.getText()) ;
            System.out.println("standardHeadRowText："+ VarUtils.standardHeadRow);
        });
        mainTabs.get("标准账单设置").add(standardHeadRowSettingPanel);

        JPanel standardNameColSettingPanel = new JPanel(new FlowLayout());
        JTextField standardNameColText = new JTextField();
        standardNameColText.setPreferredSize(new Dimension(100, 30));
        standardNameColSettingPanel.add(new JLabel("姓名所在列："));
        standardNameColSettingPanel.add(standardNameColText);
        FrameUtils.OnTextChange(standardNameColText,()->{
            VarUtils.standardNameCol = standardNameColText.getText();
            System.out.println("standardNameColText："+ VarUtils.standardNameCol);
        });
        mainTabs.get("标准账单设置").add(standardNameColSettingPanel);

        JPanel standardMoneyColSettingPanel = new JPanel(new FlowLayout());
        JTextField standardMoneyColText = new JTextField();
        standardMoneyColText.setPreferredSize(new Dimension(100, 30));
        standardMoneyColSettingPanel.add(new JLabel("金额所在列："));
        standardMoneyColSettingPanel.add(standardMoneyColText);
        FrameUtils.OnTextChange(standardMoneyColText,()->{
            VarUtils.standardMoneyCol = standardMoneyColText.getText();
            System.out.println("standardMoneyColText："+ VarUtils.standardMoneyCol);
        });
        mainTabs.get("标准账单设置").add(standardMoneyColSettingPanel);
    }
    //数据比对操作面板页面
    {
        JPanel needHandleFileCol = new JPanel();
        needHandleFileCol.add(Box.createVerticalStrut(8));
        JPanel standardFileCol = new JPanel();
        standardFileCol.add(Box.createVerticalStrut(8));
        JPanel functionalCol = new JPanel();
        functionalCol.add(Box.createVerticalStrut(8));

        needHandleFileCol.setLayout(new BoxLayout(needHandleFileCol, BoxLayout.Y_AXIS));
        standardFileCol.setLayout(new BoxLayout(standardFileCol,BoxLayout.Y_AXIS));
        functionalCol.setLayout(new BoxLayout(functionalCol,BoxLayout.Y_AXIS));

        JLabel needHandleFileListTitle = new JLabel("待处理文件列表");
        DefaultListModel<File> needHandleFileModel = new DefaultListModel<>();
        JList<File> needHandleFileList = new JList<>(needHandleFileModel);
        // 1. 正常构造：直接把 JList 当参数
        JScrollPane sp = new JScrollPane(needHandleFileList);
        sp.setPreferredSize(new Dimension(280, 200));
        sp.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        // 2. 只加标题 和 滚动面板（里面已带列表）
        needHandleFileCol.add(needHandleFileListTitle);
        needHandleFileCol.add(sp);   // *** 不要再单独 add(needHandleFileList) ***

        JLabel standardFileListTitle = new JLabel("标准文件列表");
        DefaultListModel<File> standardFileModel = new DefaultListModel<>();
        JList<File> standardFileList = new JList<>(standardFileModel);
        JScrollPane sp1 = new JScrollPane(standardFileList);
        sp1.setPreferredSize(new Dimension(280, 200));
        sp1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        standardFileCol.add(standardFileListTitle);
        standardFileCol.add(sp1);

        JButton startCmpBtn = new JButton("开始比对");
        JButton refreshListBtn = new JButton("刷新列表");
        JButton openResDirBtn = new JButton("打开结果文件夹");
        startCmpBtn.setPreferredSize(new Dimension(165,30));
        startCmpBtn.setMaximumSize(new Dimension(165,30));
        startCmpBtn.setMinimumSize(new Dimension(165,30));
        refreshListBtn.setPreferredSize(new Dimension(165,30));
        refreshListBtn.setMaximumSize(new Dimension(165,30));
        refreshListBtn.setMinimumSize(new Dimension(165,30));
        openResDirBtn.setPreferredSize(new Dimension(165,30));
        openResDirBtn.setMaximumSize(new Dimension(165,30));
        openResDirBtn.setMinimumSize(new Dimension(165,30));
        startCmpBtn.setAlignmentX(Component.CENTER_ALIGNMENT);
        refreshListBtn.setAlignmentX(Component.CENTER_ALIGNMENT);
        openResDirBtn.setAlignmentX(Component.CENTER_ALIGNMENT);


        //刷新按钮事件
        refreshListBtn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                System.out.println("刷新事件触发");
                FrameUtils.UpdateListModel(needHandleFileList,
                        needHandleFileModel,
                        GetPathUtils.scanFiles(VarUtils.needHandleFilePath,"xlsx"));
                FrameUtils.UpdateListModel(standardFileList,
                        standardFileModel,
                        GetPathUtils.scanFiles(VarUtils.standardFilePath,"xlsx"));
            }
        });
        //打开结果文件夹事件
        openResDirBtn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                File folder = new File(GetPathUtils.GetMyPath(),"result");
                try {
                    Desktop.getDesktop().open(folder);
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });
        //比对按钮事件
        startCmpBtn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                ExcelUtils.BatchCompareAndSave(
                        VarUtils.needHandleFilePath,
                        VarUtils.standardFilePath,
                        VarUtils.handleHeadRow,
                        VarUtils.handleNameCol,
                        VarUtils.handleMoneyCol,
                        VarUtils.standardHeadRow,
                        VarUtils.standardNameCol,
                        VarUtils.standardMoneyCol,
                        startCmpBtn
                );
            }
        });

        functionalCol.add(Box.createVerticalGlue());
        functionalCol.add(startCmpBtn);
        functionalCol.add(Box.createVerticalStrut(10));
        functionalCol.add(refreshListBtn);
        functionalCol.add(Box.createVerticalStrut(10));
        functionalCol.add(openResDirBtn);
        functionalCol.add(Box.createVerticalGlue());

        mainTabs.get("数据比对操作面板").setLayout(new GridLayout(1,3));
        mainTabs.get("数据比对操作面板").add(needHandleFileCol);
        mainTabs.get("数据比对操作面板").add(standardFileCol);
        mainTabs.get("数据比对操作面板").add(functionalCol);

    }
        mainFrame.setVisible(true);
    }
}
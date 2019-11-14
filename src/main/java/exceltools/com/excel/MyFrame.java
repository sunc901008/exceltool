package exceltools.com.excel;

import exceltools.com.excel.base.Common;
import exceltools.com.excel.base.Constant;
import exceltools.com.excel.handler.MergeHandler;
import exceltools.com.excel.handler.SplitHandler;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.pushingpixels.substance.api.skin.SubstanceBusinessLookAndFeel;
import org.xml.sax.SAXException;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.xml.parsers.ParserConfigurationException;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


/**
 * @author sunc
 * @date 2019/11/13 18:26
 * @description MyFrame
 */

class MyFrame extends JFrame {
    static {
        try {
            UIManager.setLookAndFeel(new SubstanceBusinessLookAndFeel());
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
    }

    MyFrame() {

        container();

        buttonListener();

        setVisible(true);
    }

    private Rectangle get(int x, int y, int with) {
        return new Rectangle(x, y, with, 20);
    }

    /**
     * 文件选择框
     */
    private JTextField selectText;
    /**
     * 选择文件按钮
     */
    private JButton jbFile;
    /**
     * 清除按钮
     */
    private JButton jbClear;
    /**
     * 合并按钮
     */
    private JButton merge;
    /**
     * 拆分按钮
     */
    private JButton split;
    /**
     * 待开发按钮
     */
    private JButton jbWrite;
    /**
     * 合并路径输出框
     */
    private JTextField destText;
    /**
     * 提示
     */
    private JLabel tips;
    /**
     * 已选择的文件/文件夹
     */
    private List<File> pathSelected = new ArrayList<>();

    /**
     * 面板容器
     */
    private void container() {

        int x1 = 10;
        int x2 = 135;
        int y1 = 10;
        int withLabel = 120;
        int withTips = 180;
        int withText = 300;
        int withButton = 100;

        setTitle("Sunc-Excel-Tool");
        setLayout(null);
        setBounds(500, 200, 720, 360);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);
        setVisible(true);

        JLabel selectLabel = new JLabel("文件路径:", SwingConstants.RIGHT);
        selectLabel.setBounds(get(x1, y1, withLabel));

        selectText = new JTextField(200);
        selectText.setBounds(get(x2, y1, withText));
        selectText.setEditable(false);

        jbFile = new JButton("选择");
        jbFile.setBounds(get(440, y1, withButton));
        jbClear = new JButton("清除");
        jbClear.setBounds(get(550, y1, withButton));

        int y3 = 60;
        merge = new JButton("合并");
        merge.setBounds(get(130, y3, withButton));

        split = new JButton("拆分");
        split.setBounds(get(240, y3, withButton));

        jbWrite = new JButton("待开发...");
        jbWrite.setBounds(get(350, y3, withButton));

        int y4 = 110;
        JLabel destLabel = new JLabel("保存位置:", SwingConstants.RIGHT);
        destLabel.setBounds(get(x1, y4, withLabel));
        destText = new JTextField(200);
        destText.setBounds(get(x2, y4, withText));
        destText.setEditable(false);

        tips = new JLabel("", SwingConstants.CENTER);
        tips.setBounds(get(200, 200, withTips));

        add(jbFile);
        add(jbClear);
        add(selectLabel);
        add(selectText);
        add(jbWrite);
        add(destLabel);
        add(destText);
        add(tips);
        add(merge);
    }

    /**
     * 选择文件按钮事件
     */
    private void jbFileListener() {
        jbFile.addActionListener(e -> {
            tips.setText("");
            JFileChooser jfc = new JFileChooser();
            jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            jfc.setCurrentDirectory(new File(Constant.BASE_PATH));
            jfc.setFileFilter(new FileNameExtensionFilter(".xlsx, .xls", "xlsx", "xls"));
            if (jfc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                selectText.setText(jfc.getSelectedFile().getAbsolutePath());
                pathSelected.clear();
                pathSelected.add(jfc.getSelectedFile());
            }
        });
    }

    /**
     * 清除按钮事件
     */
    private void jbClearListener() {
        jbClear.addActionListener(e -> {
            selectText.setText("");
            destText.setText("");
            pathSelected.clear();
            tips.setText("");
        });
    }

    /**
     * 合并按钮事件
     */
    private void mergeListener() {
        merge.addActionListener(al -> {
            tips.setForeground(Color.RED);
            tips.setText("");
            File file = pathSelected.size() > 0 ? pathSelected.get(0) : null;
            if (file == null) {
                tips.setText("请选择文件夹.");
                return;
            }

            Constant.BASE_PATH = file.getPath();
            if (!file.exists() || !file.isDirectory()) {
                tips.setText("请选择文件夹.");
                return;
            }

            String dest = file.getAbsolutePath() + File.separator + System.currentTimeMillis() + "_merge.xls";

            File[] list = file.listFiles((dir, name) -> Common.validExcelType(name));

            if (list == null || list.length <= 1) {
                tips.setText("没有需要合并的excel文件.");
                return;
            }
            List<File> files = Arrays.asList(list);
            try {
                MergeHandler.merge(dest, files);
            } catch (OpenXML4JException | ParserConfigurationException | SAXException | IOException e) {
                e.printStackTrace();
                tips.setText("执行错误.");
                return;
            }
            selectText.setText("");
            destText.setText(dest);
            tips.setForeground(Color.GREEN);
            tips.setText("合并成功.");
        });
    }

    /**
     * 拆分按钮事件
     */
    private void splitListener() {
        split.addActionListener(e -> {
            tips.setForeground(Color.RED);
            tips.setText("待开发...");

            File file = pathSelected.size() > 0 ? pathSelected.get(0) : null;
            if (file == null) {
                tips.setText("请选择excel文件.");
                return;
            }
            Constant.BASE_PATH = file.getPath();
            if (!file.isFile() || !Common.validExcelType(file)) {
                tips.setText("请选择excel文件.");
            } else {
                String dest = SplitHandler.split(file);
                if (Common.isEmpty(dest)) {
                    tips.setText("执行错误.");
                } else {
                    selectText.setText("");
                    destText.setText(dest);
                    tips.setForeground(Color.GREEN);
                    tips.setText("拆分成功.");
                }
            }
        });
    }

    /**
     * 待开发按钮事件
     */
    private void jbWriteListener() {
        jbWrite.addActionListener(e -> {
            tips.setForeground(Color.RED);
            tips.setText("待开发...");
        });
    }

    /**
     * 按钮事件绑定
     */
    private void buttonListener() {

        jbFileListener();

        jbClearListener();

        mergeListener();

        splitListener();

        jbWriteListener();
    }

}

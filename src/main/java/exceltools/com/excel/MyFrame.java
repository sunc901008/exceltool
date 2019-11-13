package exceltools.com.excel;

import exceltools.com.excel.base.Common;
import exceltools.com.excel.base.Constant;
import exceltools.com.excel.handler.MergeHandler;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;


/**
 * @author sunc
 * @date 2019/11/13 18:26
 * @description MyFrame
 */

class MyFrame extends JFrame {

    private Rectangle get(int x, int y, int with) {
        return new Rectangle(x, y, with, 20);
    }

    private JTextField selectText;
    private JButton merge;
    private JButton jbFile;
    private JButton jbClear;
    private JButton jbWrite;
    private JTextField destText;
    private JLabel tips;

    private void container() {
        setTitle("Sunc-Excel-Tool");
        setLayout(null);
        setBounds(500, 200, 700, 360);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        int x1 = 10;
        int x2 = 135;
        int y1 = 10;
        int withLabel = 120;
        int withText = 300;
        int withButton = 100;
        JLabel selectLabel = new JLabel("文件路径:", SwingConstants.RIGHT);
        selectLabel.setBounds(get(x1, y1, withLabel));

        selectText = new JTextField(200);
        selectText.setBounds(get(x2, y1, withText));
        selectText.setEditable(false);

        jbFile = new JButton("选择");
        jbFile.setBounds(get(440, y1, withButton));
        jbClear = new JButton("清除");
        jbClear.setBounds(get(530, y1, withButton));

        int y3 = 60;
        merge = new JButton("合并");
        merge.setBounds(get(130, y3, withButton));

        jbWrite = new JButton("待开发...");
        jbWrite.setBounds(get(240, y3, withButton));

        int y4 = 110;
        JLabel destLabel = new JLabel("保存位置:", SwingConstants.RIGHT);
        destLabel.setBounds(get(x1, y4, withLabel));
        destText = new JTextField(200);
        destText.setBounds(get(x2, y4, withText));
        destText.setEditable(false);

        tips = new JLabel("", SwingConstants.CENTER);
        tips.setBounds(get(240, 200, withLabel));

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

    private void buttonListener() {

        List<File> pathSelected = new ArrayList<>();
        jbFile.addActionListener(e -> {
            tips.setText("");
            JFileChooser jfc = new JFileChooser();
            jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            jfc.setCurrentDirectory(new File(Constant.BASE_PATH));
            jfc.setFileFilter(new FileNameExtensionFilter("excel文件", "xlsx", "xls", "csv"));
            if (jfc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                selectText.setText(jfc.getSelectedFile().getAbsolutePath());
                pathSelected.clear();
                pathSelected.add(jfc.getSelectedFile());
            }
        });

        jbClear.addActionListener(e -> {
            selectText.setText("");
            destText.setText("");
            pathSelected.clear();
            tips.setText("");
        });

        merge.addActionListener(e -> {
            tips.setForeground(Color.RED);
            tips.setText("");
            File file = pathSelected.size() > 0 ? pathSelected.get(0) : null;
            if (file == null) {
                tips.setText("请选择文件.");
            } else if (!Common.validExcelType(file)) {
                tips.setText("别调皮,选择excel文件.");
            } else {
                String dest = MergeHandler.merge(file);

                if (Common.isEmpty(dest)) {
                    tips.setText("未知错误.");
                } else {
                    destText.setText(dest);
                    tips.setForeground(Color.GREEN);
                    tips.setText("合并成功.");
                }
            }
        });

        jbWrite.addActionListener(e -> {
            tips.setForeground(Color.RED);
            tips.setText("待开发...");
        });
    }

    void start() {
        container();

        buttonListener();

        setVisible(true);
    }

}

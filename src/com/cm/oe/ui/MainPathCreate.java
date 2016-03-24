package com.cm.oe.ui;

import java.awt.Font;
import java.awt.Frame;
import java.awt.Label;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JTextField;

import com.cm.oe.budget.gen.BudgetWriter1;
import com.cm.oe.test.ReadExcel;

public class MainPathCreate {

	protected static final Exception NullPointerException = null;
	public JFrame frame;
	public JTextField aText = null;
	public JTextField bText = null;
	public JTextField cText = null;
	public JTextField dText = null;

	/**
	 * Launch the application.
	 */

	/**
	 * Create the application.
	 */
	public MainPathCreate() {
		initialize();
		frame.setVisible(true);
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setResizable(false);
		frame.setBounds(100, 100, 660, 425);
		frame.getContentPane().setLayout(null);

		aText = new JTextField();
		aText.setBounds(220, 39, 145, 24);
		frame.getContentPane().add(aText);

		bText = new JTextField();
		bText.setBounds(221, 102, 144, 24);
		frame.getContentPane().add(bText);

		cText = new JTextField();
		cText.setBounds(223, 160, 142, 24);
		frame.getContentPane().add(cText);

		dText = new JTextField();
		dText.setBounds(223, 220, 142, 24);
		frame.getContentPane().add(dText);

		JButton aButton = new JButton("...");
		aButton.setBounds(426, 39, 93, 23);
		aButton.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				ExcelFileFilter ef = new ExcelFileFilter();
				jFileChooser.addChoosableFileFilter(ef);
				jFileChooser.setFileFilter(ef);
				int state = jFileChooser.showOpenDialog(null);
				if (state == 1) {
					return;
				} else {
					File file = jFileChooser.getSelectedFile();
					aText.setText(file.getAbsolutePath());
				}
			}
		});
		frame.getContentPane().add(aButton);

		JButton bButton = new JButton("...");
		bButton.setBounds(426, 102, 93, 23);
		bButton.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				jFileChooser.setFileSelectionMode(1);
				int state = jFileChooser.showOpenDialog(null);
				if (state == 1) {
					return;
				} else {
					File file = jFileChooser.getSelectedFile();
					bText.setText(file.getAbsolutePath());
				}
			}
		});
		frame.getContentPane().add(bButton);

		JButton cButton = new JButton("...");
		cButton.setBounds(426, 160, 93, 23);
		cButton.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				ExcelFileFilter ef = new ExcelFileFilter();
				jFileChooser.addChoosableFileFilter(ef);
				jFileChooser.setFileFilter(ef);
				int state = jFileChooser.showOpenDialog(null);
				if (state == 1) {
					return;
				} else {
					File file = jFileChooser.getSelectedFile();
					cText.setText(file.getAbsolutePath());
				}
			}
		});
		frame.getContentPane().add(cButton);

		JButton dButton = new JButton("...");
		dButton.setBounds(426, 220, 93, 23);
		dButton.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				jFileChooser.setFileSelectionMode(1);
				int state = jFileChooser.showOpenDialog(null);
				if (state == 1) {
					return;
				} else {
					File file = jFileChooser.getSelectedFile();
					dText.setText(file.getAbsolutePath());
				}
			}
		});
		frame.getContentPane().add(dButton);

		JButton confirmButton = new JButton("生成文档");
		confirmButton.setBounds(255, 319, 178, 23);
		confirmButton.addActionListener(new ActionListener() {
			JFrame frame1 = null;
			@Override
			public void actionPerformed(ActionEvent e) throws NullPointerException {
				if (aText == null || aText.getText().equals("") || bText.getText().equals("") || bText == null
						|| cText.getText().equals("") || cText == null || dText.getText().equals("") || dText == null) {
					JOptionPane.showMessageDialog(null, "请填写全部路径");
					return;
				} else {
					int state = JOptionPane.showConfirmDialog(null, "确定选择的路径？", " ", JOptionPane.YES_NO_OPTION);
					if (state == 0) {
						try {
							boolean flag = false;
							String path1 = bText.getText();
							String path2 = cText.getText();
							String output = dText.getText() + "/";
							String tablePath = "testfiles/tables.xls";
							String excelPath = aText.getText();
							ReadB3 rB3 = new ReadB3();
							ReadExcel re = new ReadExcel(excelPath);
							Map<String, String> map = rB3.read(path1, path2);
							List<String> zhs = re.getZH();
							Set<String> keys = map.keySet();
							BudgetWriter1 xwpf = null;
							frame.setVisible(false);
							 frame1 = new MainFrame();
							Label label = new Label();
							label.setText("Please waite for a minite...");
							label.setFont(new Font("宋体", Font.PLAIN, 18));
							frame1.add(label);
							frame1.setVisible(true);
							frame1.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
							for (int i = 1; i < zhs.size(); i++) {
								String zh = zhs.get(i).trim();
								//System.out.println(keys.contains(zh));
								if (keys.contains(zh)) {
									String Path = map.get(zh);
									xwpf = new BudgetWriter1(Path, zh, path2, tablePath, excelPath, output);
									xwpf.testReadByDoc();
								} else if (!keys.contains(zhs.get(i).trim())) {
									JOptionPane.showMessageDialog(frame, "汇总表第" + (i + 1) + "行，请输入正确的站号！");
									throw NullPointerException;
								}
							}
							

							flag = true;
							if (flag) {
								frame1.setVisible(false);
								frame.setVisible(true);
								JOptionPane.showMessageDialog(frame, "生成完毕！");
							}
						} catch (NullPointerException e1) {
							frame1.setVisible(false);
							frame.setVisible(true);
							e1.printStackTrace();
							
						} catch (RuntimeException e1) {
							frame1.setVisible(false);
							JOptionPane.showMessageDialog(frame, "出错了，请选择正确的文件！");
							frame.setVisible(true);
							e1.printStackTrace();
						}catch (FileNotFoundException e2) {
							frame1.setVisible(false);
							JOptionPane.showMessageDialog(frame, "请关闭已打开的word文档！");
							frame.setVisible(true);
							e2.printStackTrace();
						} catch (Exception e1) {
							e1.printStackTrace();
						}
					} else {
						return;
					}
				}

			}
		});
		frame.getContentPane().add(confirmButton);

		JLabel lblNewLabel = new JLabel("一体化基站勘察汇总表\r\n");
		lblNewLabel.setBounds(57, 43, 132, 15);
		frame.getContentPane().add(lblNewLabel);

		JLabel lblg = new JLabel("4G工程基站预算表路径");
		lblg.setBounds(57, 106, 132, 15);
		frame.getContentPane().add(lblg);

		JLabel lblgg = new JLabel("3G4G工程基站预算表路径");
		lblgg.setBounds(57, 164, 132, 15);
		frame.getContentPane().add(lblgg);

		JLabel label_2 = new JLabel("文件生成路径");
		label_2.setBounds(57, 224, 121, 15);
		frame.getContentPane().add(label_2);

	}
}

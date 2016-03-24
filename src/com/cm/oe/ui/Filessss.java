package com.cm.oe.ui;

import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.ButtonGroup;
import javax.swing.JButton;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JRadioButton;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.cm.oe.test.MatcherValitation;

import javax.swing.SwingConstants;
import java.awt.Color;
import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;

public class Filessss {
	private JFrame frame;

	/**
	 * Launch the application.
	 ****/


	/**
	 * Create the application.
	 */
	public Filessss() {
		initialize();
		this.frame.setVisible(true);
		
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new MainFrame();
		frame.setResizable(false);
		frame.setResizable(false);
		frame.setTitle("标题");
		frame.setBounds(100, 100, 481, 297);
		frame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);		
		frame.getContentPane().setLayout(null);
		final TextField text=new TextField();
		text.setBounds(127, 70, 160, 23);
		
		
		
		text.setText("C:\\Users\\admin\\Desktop");
		frame.getContentPane().add(text);
		
		JButton jButton=new JButton();
		jButton.setBounds(321, 70, 51, 23);
		jButton.setBackground(Color.WHITE);
		jButton.setForeground(Color.DARK_GRAY);
		jButton.setText("...");
		jButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				        JFileChooser jFileChooser=new JFileChooser();
			            jFileChooser.setFileSelectionMode(1);  
			            int state = jFileChooser.showOpenDialog(null); 
			            if (state == 1) {  
			                return;  
			            } else {  
			                File file = jFileChooser.getSelectedFile(); 
			                text.setText(file.getAbsolutePath());  
			            }   
				
			}
		});
		frame.getContentPane().add(jButton);
		
		
		//为导出表格添加事件
		JButton btnNewButton = new JButton("开始校验");
		btnNewButton.setBounds(159, 131, 177, 23);
		btnNewButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				MatcherValitation mv = new MatcherValitation();
				try {
					mv.files(text.getText());
				} catch (Exception e1) {
					e1.printStackTrace();
				}
		}});
		frame.getContentPane().add(btnNewButton);
		
		JLabel lblNewLabel = new JLabel("文件路径");
		lblNewLabel.setBounds(43, 37, 54, 15);
		frame.getContentPane().add(lblNewLabel);
		

		
		
	
	}
}

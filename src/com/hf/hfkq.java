package com.hf;

import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.Map; 
import java.applet.*;
import javax.swing.*;
import javax.swing.GroupLayout.Alignment;
import javax.swing.filechooser.FileSystemView;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import com.eltima.components.ui.DatePicker;
import java.util.Date;
import java.util.Locale;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.hf.ZkemSDK;
import com.hf.ConfigReader;

public class hfkq extends JFrame{
	//定义该图形中所需的组件的引用。  
    private JFrame f;  
    private Button btnConn;
    private Button btnChoose;
    private Button btnDownload;
    private Button btnUsers;
    private Button btnQuit;
    
    private JLabel jl;
    private TextField tf;
    
    private JLabel jl_kq;
    private DatePicker datepick;
    
    private JComboBox comboBox;
    private Button btnGetResult;
    private Button btnAnalyseLog;
    
    private JLabel jl_status;
    private Dialog dlgError;
    
    private String DownloadPath="";
    
    
	public static void main(String[] args)  
    {  
        new hfkq();  
    }
	
	public hfkq() {
		init();
		
	}
	
	
        
      
    public void init()  
    {  
        f = new JFrame("考勤记录处理软件");  
          
        //对frame进行基本设置  
        f.setBounds(300, 100, 600, 500);  
        //f.setDefaultCloseOperation(EXIT_ON_CLOSE);
        f.setResizable(false);
        //f.setLayout(new FlowLayout());  
        JPanel panel = new JPanel();
        //初始化组件并加到frame中
        
        //jl_title.setHorizontalAlignment(JLabel.CENTER);
        jl = new JLabel("软件操作目录：");
        tf = new TextField(30);// 创建单行文本对象60长度大小字符
        String defaultPath = System.getProperty("user.dir");
        tf.setText(defaultPath);
        DownloadPath=defaultPath;
        tf.setEditable(false);
        
        btnChoose = new Button("选择目录");  
        btnConn = new Button("测试考勤机连接");  
        btnDownload = new Button("下载全部考勤记录");  
        btnUsers = new Button("更新所有考勤人员");
        jl_kq = new JLabel("考勤日期区间选择：");
        datepick = getDatePicker();
        
        comboBox = new JComboBox();
        comboBox.addItem("签到");
        comboBox.addItem("签退");
        btnGetResult = new Button("获取指定时间区间考勤记录");
        btnAnalyseLog = new Button("分析考勤数据");
        btnAnalyseLog.setEnabled(false);
        btnQuit = new Button("退出");
        
        jl_status = new JLabel("");
        
        //为指定的Container创建GroupLayout
        GroupLayout layout = new GroupLayout(panel);
        layout.setAutoCreateGaps(true);
        layout.setAutoCreateContainerGaps(true);
        panel.setLayout(layout);
        //创建GroupLayout的水平连续组，越先加入的ParalleGroupLayout优先级越高
        GroupLayout.SequentialGroup hGroup = layout.createSequentialGroup();
        hGroup.addGroup(layout.createParallelGroup().addComponent(jl).addComponent(btnConn).addComponent(jl_kq).addComponent(btnGetResult));        
        hGroup.addGroup(layout.createParallelGroup().addComponent(tf).addComponent(btnDownload).addComponent(datepick).addComponent(btnAnalyseLog).addComponent(jl_status));        
        hGroup.addGroup(layout.createParallelGroup().addComponent(btnChoose).addComponent(btnUsers).addComponent(comboBox).addComponent(btnQuit));
        
        layout.setHorizontalGroup(hGroup);
        //创建GroupLayout的垂直连续组，越线加入的ParallelGroup，优先级越高
        GroupLayout.SequentialGroup vGroup = layout.createSequentialGroup();
        vGroup.addGroup(layout.createParallelGroup().addComponent(jl).addComponent(tf).addComponent(btnChoose));
        vGroup.addGroup(layout.createParallelGroup().addComponent(btnConn).addComponent(btnDownload).addComponent(btnUsers));
        vGroup.addGroup(layout.createParallelGroup().addComponent(jl_kq).addComponent(datepick).addComponent(comboBox));
        vGroup.addGroup(layout.createParallelGroup().addComponent(btnGetResult).addComponent(btnAnalyseLog).addComponent(btnQuit)); 
        vGroup.addGroup(layout.createParallelGroup().addComponent(jl_status));
        //设置垂直组
        layout.setVerticalGroup(vGroup);
        //加载一下窗体上的事件  
        myEvent();  
        
        f.setContentPane(panel);
        f.pack();
        f.setLocationRelativeTo(null);
        //显示窗口  
        f.setVisible(true);  
    }  
      
    //事件响应  
    private void myEvent()  
    {  
        f.addWindowListener(new WindowAdapter(){  
            public void windowClosing(WindowEvent e)  
            {  
                System.exit(0);  
            }  
        }); 
        //测试连接  
        btnConn.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub 
        		btnConn.setEnabled(false);
        		String iniPath = System.getProperty("user.dir");
        		//JOptionPane.showMessageDialog(null, "当前路径是："+iniPath, "提示", JOptionPane.INFORMATION_MESSAGE); 
        		try{
        			ZkemSDK sdk = new ZkemSDK();
            		boolean  connFlag = true;  
            		
            		String strError="";
            		
            		ConfigReader cr = new  ConfigReader(iniPath+"\\config.ini");
                    System.out.println(cr.get("config", "IP"));
                    
                    String strIPs = cr.get("config", "IP").get(0);
                    String[] strIP = strIPs.split(",");
                    for(int i=0;i<strIP.length;i++){
                    	System.out.println(strIP[i]);
                    	if(!sdk.connect(strIP[i], 4370)){
                    		connFlag = false;
                    		strError = strError + strIP[i]+" ";
                    	}else{
                    		sdk.disConnect();
                    	}
                    }
                    if(connFlag){
                    	JOptionPane.showMessageDialog(null, "所有考勤机连接成功！", "提示", JOptionPane.INFORMATION_MESSAGE);
                    	btnConn.setEnabled(true);
                    }else{
                    	JOptionPane.showMessageDialog(null, strError+"考勤机连接失败！", "错误", JOptionPane.ERROR_MESSAGE);
                    	btnConn.setEnabled(true);
                    }
        		}catch(Exception err){
        			JOptionPane.showMessageDialog(null, err.toString(), "错误", JOptionPane.ERROR_MESSAGE);
        			btnConn.setEnabled(true);
        		}
        	}  
        }); 
        //选择保存目录
        btnChoose.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub
        		String iniPath = System.getProperty("user.dir");
        		JFileChooser fileChooser = new JFileChooser(iniPath);
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnVal = fileChooser.showOpenDialog(fileChooser);
                if(returnVal == JFileChooser.APPROVE_OPTION){       
                	DownloadPath= fileChooser.getSelectedFile().getAbsolutePath();//这个就是你选择的文件夹的路径
                    System.out.println(DownloadPath);
                    tf.setText(DownloadPath);
                }
                
        	}  
        });
        //下载考勤记录
        btnDownload.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub  
        		if(DownloadPath.equals("")){
        			JOptionPane.showMessageDialog(null, "未选择操作目录，不知下载到何处！", "提示", JOptionPane.INFORMATION_MESSAGE); 
        		}else{
        			btnDownload.setEnabled(false);
        			System.out.println("下载考勤记录到："+DownloadPath);
        			ZkemSDK sdk = new ZkemSDK();             		
            		String strError=""; 
            		String iniPath = System.getProperty("user.dir");
            		ConfigReader cr = new  ConfigReader(iniPath+"\\config.ini");
                    System.out.println(cr.get("config", "IP"));
                    String strIPs = cr.get("config", "IP").get(0);
                    String[] strIP = strIPs.split(",");
                                     
                    int iMachineNumber=0;
                    for(int i=0;i<strIP.length;i++){
                    	System.out.println(strIP[i]);
                    	if(!sdk.connect(strIP[i], 4370)){
                    		strError = strError + strIP[i]+" ";
                    	}else{
                    		sdk.readGeneralLogData(iMachineNumber);
                    		if("\\".equals(DownloadPath.substring(DownloadPath.length()-1))){
                    			saveFileItemsToTxt(iMachineNumber,sdk.getGeneralLogData(),DownloadPath+"output.txt");
                    		}else{
                    			saveFileItemsToTxt(iMachineNumber,sdk.getGeneralLogData(),DownloadPath+"\\output.txt");
                    		}             		
                    	}
                    	iMachineNumber++;
                    }
                    if(strError.equals("")){
                    	JOptionPane.showMessageDialog(null, "下载完成！", "提示", JOptionPane.INFORMATION_MESSAGE); 
                    	btnDownload.setEnabled(true);
                    }else{
                    	JOptionPane.showMessageDialog(null, strError+"考勤数据下载失败！", "提示", JOptionPane.ERROR_MESSAGE); 
                    	btnDownload.setEnabled(true);
                    }
                    
        		}
        	}  
        });
        //获取所有考勤人员
        btnUsers.addActionListener(new ActionListener() {
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub  
        		btnUsers.setEnabled(false);
        		System.out.println("下载人员信息到："+DownloadPath);
        		ZkemSDK sdk = new ZkemSDK();             		
            	String strError=""; 
            	String iniPath = System.getProperty("user.dir");
            	ConfigReader cr = new  ConfigReader(iniPath+"\\config.ini");
                System.out.println(cr.get("config", "IP"));
                String strIPs = cr.get("config", "IP").get(0);
                String[] strIP = strIPs.split(",");
                
                System.out.println(strIP[0]);
                if(!sdk.connect(strIP[0], 4370)){
                    strError = strError + strIP[0]+" ";
                }else{
                    sdk.getUserInfo();
                    downUsers(sdk.getUserInfo());
                }
                    
                if(strError.equals("")){
                    JOptionPane.showMessageDialog(null, "更新完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
                    btnUsers.setEnabled(true);
                }else{
                    JOptionPane.showMessageDialog(null, strError+"更新人员信息失败！", "提示", JOptionPane.ERROR_MESSAGE); 
                    btnUsers.setEnabled(true);
                }
        	}
        });
        //获取指定时间区间考勤记录
        btnGetResult.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub  
        		btnGetResult.setEnabled(false);
        		String strDT = datepick.getText();
        		String strAP = comboBox.getSelectedItem().toString();
        		System.out.print(strDT+" "+strAP);
        		
        		ZkemSDK sdk = new ZkemSDK();             		
        		String strError=""; 
        		String iniPath = System.getProperty("user.dir");
        		ConfigReader cr = new  ConfigReader(iniPath+"\\config.ini");
                System.out.println(cr.get("config", "IP"));
                String strIPs = cr.get("config", "IP").get(0);
                String[] strIP = strIPs.split(",");
                
                //清空之前下载的考勤记录
                String strurl1 = "jdbc:Access:///"+iniPath+"/users.mdb";
            	Connection conn1 = null;
            	String strsql1 = "";
            	try{
	            	Class.forName("com.hxtt.sql.access.AccessDriver");
	                conn1 = DriverManager.getConnection(strurl1);
	                strsql1 = "delete from LogTmp;";
	                PreparedStatement ps1 = conn1.prepareStatement(strsql1);
	                ps1.execute();
	            }catch(Exception e1){
	            	e1.printStackTrace();  
	            }
                
                int iMachineNumber=0;
                for(int i=0;i<strIP.length;i++){
                	System.out.println(strIP[i]);
                	if(!sdk.connect(strIP[i], 4370)){
                		strError = strError + strIP[i]+" ";
                	}else{
                		sdk.readGeneralLogData(iMachineNumber);
                		downLogs(iMachineNumber,sdk.getGeneralLogData(),strDT,strAP);
                	}
                	iMachineNumber++;
                }
                if(strError.equals("")){
                    JOptionPane.showMessageDialog(null, "获取"+strDT+strAP+"记录成功！", "提示", JOptionPane.INFORMATION_MESSAGE);
                    btnAnalyseLog.setEnabled(true);
                    btnGetResult.setEnabled(true);
                }else{
                    JOptionPane.showMessageDialog(null, strError+"获取考勤记录失败！", "提示", JOptionPane.ERROR_MESSAGE); 
                    btnAnalyseLog.setEnabled(true);
                    btnGetResult.setEnabled(true);
                }
        	}
        });
        //分析考勤记录 ---未打卡、迟到、早退、加班等
        btnAnalyseLog.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub  
        		if(DownloadPath.equals("")){
        			JOptionPane.showMessageDialog(null, "未选择操作目录，分析结果无法输出！", "提示", JOptionPane.INFORMATION_MESSAGE); 
        		}else{
        			String strDT = datepick.getText();
        			String strAP = comboBox.getSelectedItem().toString();
        			if("\\".equals(DownloadPath.substring(DownloadPath.length()-1))){
        				AnalyseLog(strDT,strAP,DownloadPath+strDT+strAP+"分析结果.xls");
            		}else{
            			AnalyseLog(strDT,strAP,DownloadPath+"\\"+strDT+strAP+"分析结果.xls");
            		}      			
        			btnAnalyseLog.setEnabled(false);
        		}
        	}
        });
        //退出程序
        btnQuit.addActionListener(new ActionListener() {  
        	@Override  
        	public void actionPerformed(ActionEvent e) {  
        	// TODO Auto-generated method stub  
        		Object[] options = { "退出程序", "取消" }; 
        		int result = JOptionPane.showOptionDialog(null, "确定退出？", "Warning", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null, options, options[0]); 
        		if(result == 0){
        			f.setVisible(false);
                    f.dispose();
                    System.exit(0); 
        		}          
        	}  
        }); 
    }
    
    public static void saveFileItemsToTxt(int machineNo,List<Map<String,Object>> list,String strFileName){  
        
        OutputStreamWriter outFile = null;  
        FileOutputStream fileName;  
        String strItems = null;  
        try{
        	//允许追加存储
            fileName = new FileOutputStream(strFileName,true); 
            //不允许追加存储
            //fileName = new FileOutputStream(strFileName); 
              
            outFile = new OutputStreamWriter(fileName);  
            
            for(Map<String,Object> map:list){  
                for (String key : map.keySet()) { 
                    strItems = Integer.toString(machineNo)+"|"+map.get("EnrollNumber")+"|"+map.get("Time")+"|"+map.get("VerifyMode")+"|"+map.get("InOutMode")+"\n";
                }  
                outFile.write(strItems);  
            } 
        }  
        catch(Exception e)  
        {  
            e.printStackTrace();  
        }finally{  
            try {  
                outFile.flush();  
                outFile.close();  
            } catch (IOException e) {  
                // TODO Auto-generated catch block  
                e.printStackTrace();  
            }  
        }  
    }
    
    //保存考勤人员信息
    public static void downUsers(List<Map<String,Object>> list){ 
    	String iniPath = System.getProperty("user.dir");
        String strurl = "jdbc:Access:///"+iniPath+"/users.mdb";
    	Connection conn = null;
    	//Statement stmt = null;
    	String strsql = "";
    	//ResultSet rs = null;
        try{
            Class.forName("com.hxtt.sql.access.AccessDriver");
            conn = DriverManager.getConnection(strurl);
            strsql = "delete from Users;";
            //stmt = conn.createStatement();
            //stmt.execute("delete from Users");
            for(Map<String,Object> map:list){
            	//System.out.print(map.get("Enabled").toString());
            	if("true".equals(map.get("Enabled").toString())){
            		String abc = "insert into Users(EnrollNumber,UserName,Privilege,Enabled) values('"+map.get("EnrollNumber")+"','"+map.get("Name")+"','"+map.get("Privilege")+"',"+map.get("Enabled")+");";
                    strsql+=abc;
                    //stmt.execute(abc);
            	}
            } 
            PreparedStatement ps = conn.prepareStatement(strsql);
            ps.execute();
        }  
        catch(Exception e)  
        {  
            e.printStackTrace();  
        }
    }
    
    //保存指定日期的考勤记录
    public static void downLogs(int machineNo,List<Map<String,Object>> list,String strDT,String strAP){ 
    	String iniPath = System.getProperty("user.dir");
        String strurl = "jdbc:Access:///"+iniPath+"/users.mdb";
    	Connection conn = null;
    	//Statement stmt = null;
    	//ResultSet rs = null;
    	String strsql = "";
    	DateFormat df = new SimpleDateFormat("yyyy-m-d HH:mm:ss");
        try{
            Class.forName("com.hxtt.sql.access.AccessDriver");
            conn = DriverManager.getConnection(strurl);
            //stmt = conn.createStatement();
            //stmt.execute("delete from LogTmp");
            //strsql = "delete from LogTmp;";
            for(Map<String,Object> map:list){
            	if("签到".equals(strAP)){
	            	if(strDT.equals(map.get("Time").toString().substring(0,strDT.length()))&&(df.parse(map.get("Time").toString()).compareTo(df.parse(strDT+" 12:00:00")))<0){
	            		String abc = "insert into LogTmp(machineNo,EnrollNumber,LogTime,VerifyMode,InOutMode) values('"+Integer.toString(machineNo)+"','"+map.get("EnrollNumber")+"','"+map.get("Time")+"','"+map.get("VerifyMode")+"',"+map.get("InOutMode")+");";
	            		//stmt.execute(abc);
	            		strsql+=abc;
	            	}
            	}else if("签退".equals(strAP)){
            		if(strDT.equals(map.get("Time").toString().substring(0,strDT.length()))&&(df.parse(map.get("Time").toString()).compareTo(df.parse(strDT+" 12:00:00")))>0){
	            		String abc = "insert into LogTmp(machineNo,EnrollNumber,LogTime,VerifyMode,InOutMode) values('"+Integer.toString(machineNo)+"','"+map.get("EnrollNumber")+"','"+map.get("Time")+"','"+map.get("VerifyMode")+"',"+map.get("InOutMode")+");";
	            		//stmt.execute(abc);
	            		strsql+=abc;
	            	}
            	}
            }
            PreparedStatement ps = conn.prepareStatement(strsql);
            ps.execute();
        }  
        catch(Exception e)  
        {  
            e.printStackTrace();  
        }
    }
    
    //分析考勤记录
    public static void AnalyseLog(String strDT,String strAP,String strFileName){ 
    	String iniPath = System.getProperty("user.dir");
        String strurl = "jdbc:Access:///"+iniPath+"/users.mdb";
    	Connection conn = null;
    	Statement stmt = null;
    	ResultSet rs = null;
    	DateFormat df = new SimpleDateFormat("yyyy-m-d HH:mm:ss");
    	//OutputStreamWriter outFile = null;  
        //FileOutputStream fileName;  
        //String strItems = null;  
        try{
            Class.forName("com.hxtt.sql.access.AccessDriver");
            conn = DriverManager.getConnection(strurl);
            stmt = conn.createStatement();
            String strsql = "select a.EnrollNumber,a.UserName from Users a left join LogTmp b on a.EnrollNumber=b.EnrollNumber where b.LogTime is null";
            rs=stmt.executeQuery(strsql);
            //新建xls文件
            WritableWorkbook workbook = Workbook.createWorkbook(new File(strFileName)); 
            //新建工作表保存未打卡人员
            WritableSheet firstSheet = workbook.createSheet("未打卡人员", 0);
            Label label1 = new Label(0,0,"考勤号");
            Label label2 = new Label(1,0,"姓名");
            firstSheet.addCell(label1);
            firstSheet.addCell(label2);
            Label label3;
            Label label4;
            int i = 1;
            while(rs.next()){
            	label3 = new Label(0,i,rs.getString(1));
            	label4 = new Label(1,i,rs.getString(2));
            	firstSheet.addCell(label3);
                firstSheet.addCell(label4);
                i++;
        	}
            strsql  = "select a.EnrollNumber,a.UserName,b.LogTime from Users a left join LogTmp b on a.EnrollNumber=b.EnrollNumber where b.LogTime is not null order by b.LogTime";
            rs = null;
            rs = stmt.executeQuery(strsql);
            WritableSheet secondSheet = workbook.createSheet("迟到or加班名单", 0);
            Label label5 = new Label(0,0,"考勤号");
            Label label6 = new Label(1,0,"姓名");
            Label label7 = new Label(2,0,"打卡时间");
            secondSheet.addCell(label5);
            secondSheet.addCell(label6);
            secondSheet.addCell(label7);
            Label label8;
            Label label9;
            Label label10;
            i=1;
            if("签到".equals(strAP)){
            	while(rs.next()){
            		if((df.parse(rs.getString(3)).compareTo(df.parse(strDT+" 9:00:00")))>0){
	                	label8 = new Label(0,i,rs.getString(1));
	                	label9 = new Label(1,i,rs.getString(2));
	                	label10 = new Label(2,i,rs.getString(3));
	                	secondSheet.addCell(label8);
	                    secondSheet.addCell(label9);
	                    secondSheet.addCell(label10);
	                    i++;
            		}
                }
            }else if("签退".equals(strAP)){
            	while(rs.next()){
            		if((df.parse(rs.getString(3)).compareTo(df.parse(strDT+" 18:30:00")))>0){
	                	label8 = new Label(0,i,rs.getString(1));
	                	label9 = new Label(1,i,rs.getString(2));
	                	label10 = new Label(2,i,rs.getString(3));
	                	secondSheet.addCell(label8);
	                    secondSheet.addCell(label9);
	                    secondSheet.addCell(label10);
	                    i++;
            		}
                }
            }           
            //打开流，开始写文件
            workbook.write();
            //关闭流
            workbook.close();
        }  
        catch(Exception e)  
        {  
        	JOptionPane.showMessageDialog(null, "分析结果导出失败！", "提示", JOptionPane.ERROR_MESSAGE); 
            e.printStackTrace();
            return;
        } 
        JOptionPane.showMessageDialog(null, "分析结果已输出至目录"+strFileName, "提示", JOptionPane.INFORMATION_MESSAGE); 
    }
    
    private static DatePicker getDatePicker() {
        final DatePicker datepick;
        // 格式
        String DefaultFormat = "yyyy-M-d";
        // 当前时间
        Date date = new Date();
        // 字体
        Font font = new Font("Times New Roman", Font.BOLD, 14);

        Dimension dimension = new Dimension(177, 24);
        //构造方法（初始时间，时间显示格式，字体，控件大小）
        datepick = new DatePicker(date, DefaultFormat, font, dimension);   
        datepick.setLocation(137, 83);//设置起始位置
        // 设置国家
        datepick.setLocale(Locale.CHINA);
        // 设置时钟面板可见
        datepick.setTimePanleVisible(false);
        return datepick;
    }
}
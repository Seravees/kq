package com.qqq.main;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.qqq.data.Dao;

public class Main {

	/**
	 * @param args
	 * @throws IOException
	 * @throws ParseException
	 */
	public static void main(String[] args) throws IOException, ParseException {
		// TODO Auto-generated method stub
		Date start = new Date();

		String path = System.getProperty("user.dir");
		@SuppressWarnings("resource")
		BufferedReader r = new BufferedReader(new InputStreamReader(
				new FileInputStream(new File(path + "/config/config.txt")),
				"UTF-8"));
		String str;
		List<String> list = new ArrayList<String>();
		while ((str = r.readLine()) != null) {
			System.out.println(str);
			list.add(str);
		}

		for (String s : list) {
			if (s.startsWith("﻿源路径：")) {
				Dao.setInPath(s.substring(new String("﻿源路径：").length()));
			}
			if (s.startsWith("目标路径：")) {
				Dao.setOutPath(s.substring(new String("目标路径：").length()));
			}
			if (s.startsWith("源表格：")) {
				Dao.setSrc(s.substring(new String("源表格：").length()));
			}
			if (s.startsWith("人员名单：")) {
				Dao.setName(s.substring(new String("人员名单：").length()));
			}
			if (s.startsWith("排班1：")) {
				Dao.setPb1(s.substring(new String("排班1：").length()));
			}
			if (s.startsWith("排班2：")) {
				Dao.setPb2(s.substring(new String("排班2：").length()));
			}
			if (s.startsWith("考勤打卡1：")) {
				Dao.setKq1(s.substring(new String("考勤打卡1：").length()));
			}
			if (s.startsWith("考勤打卡2：")) {
				Dao.setKq2(s.substring(new String("考勤打卡2：").length()));
			}
			if (s.startsWith("请假：")) {
				Dao.setHoliday(s.substring(new String("请假：").length()));
			}
			if (s.startsWith("统筹加班：")) {
				Dao.setAdd1(s.substring(new String("统筹加班：").length()));
			}
			if (s.startsWith("临时加班：")) {
				Dao.setAdd2(s.substring(new String("临时加班：").length()));
			}
			if (s.startsWith("临时加班未结束：")) {
				Dao.setAdd2se(s.substring(new String("临时加班未结束：").length()));
			}
			if (s.startsWith("外出：")) {
				Dao.setOut(s.substring(new String("外出：").length()));
			}
			// if (s.startsWith("开始日期：")) {
			// startDate = s.substring(new String("开始日期：").length());
			// }
		}

		System.out.println(Dao.getInPath());

		Dao.createPerson();
		Dao.setPBs();
		// Dao.setPBs(startDate);
		Dao.setKQs();
		// Dao.setKQs(startDate);
		Dao.setOuts();
		Dao.setHolidays();
		// Dao.setHolidays(startDate);
		// Dao.sumPBs();
		Dao.setAdds();
		Dao.fix();
		Dao.merge();

		Date end = new Date();
		System.out.println("use" + (end.getTime() - start.getTime()) / 1000
				+ "s");

	}
}

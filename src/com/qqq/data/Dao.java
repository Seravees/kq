package com.qqq.data;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.qqq.model.Add;
import com.qqq.model.Holiday;
import com.qqq.model.HolidayPerson;
import com.qqq.model.KQ;
import com.qqq.model.Out;
import com.qqq.model.PB;
import com.qqq.model.Person;
import com.qqq.util.Tools;

public class Dao {
	// static String inPath = "C:/Users/qqq/Desktop/src/";
	// static String outPath = "C:/Users/qqq/Desktop/test/";
	// static String src = "个人考勤一览表";
	// static String name = "人员名单";
	// static String pb1 = "人员排班情况表3.26-3.31";
	// static String pb2 = "人员排班情况表4.1-4.25";
	// static String kq1 = "考勤日报汇总表3.26-3.31";
	// static String kq2 = "考勤日报汇总表4.1-4.25";
	// static String holiday = "请假查询_1";
	// static String add1 = "统筹加班_1";
	// static String add2 = "临时加班_1";
	// static String add2se = "临时加班_1-未结束";
	// static String out = "外出申请_1";
	static String XLS = "xls";
	static String XLSX = "xlsx";

	static String inPath;
	static String outPath;
	static String src;
	static String name;
	static String pb1;
	static String pb2;
	static String kq1;
	static String kq2;
	static String holiday;
	static String add1;
	static String add2;
	static String add2se;
	static String out;

	public static String getInPath() {
		return inPath;
	}

	public static void setInPath(String inPath) {
		Dao.inPath = inPath;
	}

	public static String getOutPath() {
		return outPath;
	}

	public static void setOutPath(String outPath) {
		Dao.outPath = outPath;
	}

	public static String getSrc() {
		return src;
	}

	public static void setSrc(String src) {
		Dao.src = src;
	}

	public static String getName() {
		return name;
	}

	public static void setName(String name) {
		Dao.name = name;
	}

	public static String getPb1() {
		return pb1;
	}

	public static void setPb1(String pb1) {
		Dao.pb1 = pb1;
	}

	public static String getPb2() {
		return pb2;
	}

	public static void setPb2(String pb2) {
		Dao.pb2 = pb2;
	}

	public static String getKq1() {
		return kq1;
	}

	public static void setKq1(String kq1) {
		Dao.kq1 = kq1;
	}

	public static String getKq2() {
		return kq2;
	}

	public static void setKq2(String kq2) {
		Dao.kq2 = kq2;
	}

	public static String getAdd1() {
		return add1;
	}

	public static void setAdd1(String add1) {
		Dao.add1 = add1;
	}

	public static String getAdd2() {
		return add2;
	}

	public static void setAdd2(String add2) {
		Dao.add2 = add2;
	}

	public static String getAdd2se() {
		return add2se;
	}

	public static void setAdd2se(String add2se) {
		Dao.add2se = add2se;
	}

	public static String getOut() {
		return out;
	}

	public static void setOut(String out) {
		Dao.out = out;
	}

	public static String getHoliday() {
		return holiday;
	}

	public static void setHoliday(String holiday) {
		Dao.holiday = holiday;
	}

	// TODO Auto-generated method stub
	public static void createPerson() throws IOException {
		List<List<Object>> names = Tools.readAll(inPath, name, XLSX);
		List<Person> persons = new ArrayList<Person>();
		for (int i = 1; i < names.size(); i++) {
			Person person = new Person();
			person.setName((String) names.get(i).get(1));
			person.setDepartment((String) names.get(i).get(0));
			persons.add(person);
		}
		for (Person person : persons) {
			String fileName = person.getDepartment() + "-" + person.getName();
			Tools.writerString(inPath, src, outPath, fileName, XLS, 1, 1,
					person.getName(), false);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS, 1, 5,
					person.getDepartment(), false);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS, 1,
					10, new SimpleDateFormat("MM").format(new Date()), false);
		}
	}

	public static void setPBs(String date) throws IOException, ParseException {
		List<List<Object>> pb1s = Tools.readAll(inPath, pb1, XLS);
		List<List<Object>> pb2s = Tools.readAll(inPath, pb2, XLS);

		List<PB> pbs = new ArrayList<PB>();
		for (int i = 1; i < pb1s.size(); i++) {
			for (int j = 3; j < pb1s.get(i).size(); j++) {
				PB pb = new PB();
				pb.setName((String) pb1s.get(i).get(2));
				pb.setDepartment((String) pb1s.get(i).get(0));
				String temp = (String) pb1s.get(0).get(j);
				pb.setDate(temp.substring(0, 5));
				pb.setWeekday(temp.substring(6, 9));
				pb.setPb((String) pb1s.get(i).get(j));
				pbs.add(pb);
			}
		}
		for (int i = 1; i < pb2s.size(); i++) {
			for (int j = 3; j < pb2s.get(i).size(); j++) {
				PB pb = new PB();
				pb.setName((String) pb2s.get(i).get(2));
				pb.setDepartment((String) pb2s.get(i).get(0));
				String temp = (String) pb2s.get(0).get(j);
				pb.setDate(temp.substring(0, 5));
				pb.setWeekday(temp.substring(6, 9));
				pb.setPb((String) pb2s.get(i).get(j));
				pbs.add(pb);
			}
		}

		SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
		Date start = sdf.parse(date);

		for (PB pb : pbs) {
			String fileName = pb.getDepartment() + "-" + pb.getName();
			Date end = sdf.parse(pb.getDate());
			long diff = end.getTime() - start.getTime();
			int rowNum = (int) (diff / 1000 / 60 / 60 / 24 + 5);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS,
					rowNum, 0, pb.getDate(), true);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS,
					rowNum, 1, pb.getWeekday(), true);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS,
					rowNum, 2, pb.getPb(), true);

		}
	}

	public static void setKQs(String date) throws IOException, ParseException {
		List<List<Object>> kq1s = Tools.readAll(inPath, kq1, XLS);
		List<List<Object>> kq2s = Tools.readAll(inPath, kq2, XLS);

		List<KQ> kqs = new ArrayList<KQ>();

		for (int i = 2; i < kq1s.size(); i++) {
			for (int j = 1; j < kq1s.get(i).size() / 2; j++) {
				KQ kq = new KQ();
				kq.setName((String) kq1s.get(i).get(1));
				kq.setDepartment((String) kq1s.get(i).get(0));
				String temp = (String) kq1s.get(0).get(j * 2);
				kq.setDate(temp.substring(0, 5));
				kq.setWeekday(temp.substring(6, 9));
				kq.setStart((String) kq1s.get(i).get(j * 2));
				kq.setEnd((String) kq1s.get(i).get(j * 2 + 1));
				kqs.add(kq);
			}
		}
		for (int i = 2; i < kq2s.size(); i++) {
			for (int j = 1; j < kq2s.get(i).size() / 2; j++) {
				KQ kq = new KQ();
				kq.setName((String) kq2s.get(i).get(1));
				kq.setDepartment((String) kq2s.get(i).get(0));
				String temp = (String) kq2s.get(0).get(j * 2);
				kq.setDate(temp.substring(0, 5));
				kq.setWeekday(temp.substring(6, 9));
				kq.setStart((String) kq2s.get(i).get(j * 2));
				kq.setEnd((String) kq2s.get(i).get(j * 2 + 1));
				kqs.add(kq);
			}
		}

		SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
		Date start = sdf.parse(date);

		for (KQ kq : kqs) {
			String fileName = kq.getDepartment() + "-" + kq.getName();
			Date end = sdf.parse(kq.getDate());
			long diff = end.getTime() - start.getTime();
			int rowNum = (int) (diff / 1000 / 60 / 60 / 24 + 5);

			Tools.writerString(outPath, fileName, outPath, fileName, XLS,
					rowNum, 3, kq.getStart(), true);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS,
					rowNum, 4, kq.getEnd(), true);

		}

	}

	public static void setHolidays(String date) throws IOException,
			ParseException {
		List<List<Object>> holidays = Tools.readAll(inPath, holiday, XLS);
		List<Holiday> hols = new ArrayList<Holiday>();

		for (int i = 3; i < holidays.size(); i++) {
			Holiday hol = new Holiday();
			hol.setDate(((String) holidays.get(i).get(3)).substring(5, 10));
			hol.setDepartment((String) holidays.get(i).get(2));
			hol.setName((String) holidays.get(i).get(1));
			hol.setStart(((String) holidays.get(i).get(3)).substring(5));
			hol.setEnd(((String) holidays.get(i).get(4)).substring(5));
			hol.setType((String) holidays.get(i).get(8));
			double hours = ((Double) holidays.get(i).get(10) * 8 + (Double) holidays
					.get(i).get(11));
			hol.setHours(hours);
			hols.add(hol);
		}

		List<HolidayPerson> holidayPersons = new ArrayList<HolidayPerson>();
		for (Holiday holiday : hols) {

			if (holidayPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < holidayPersons.size(); i++) {
					if (holidayPersons.get(i).getName()
							.equals(holiday.getName())
							&& holidayPersons.get(i).getDepartment()
									.equals(holiday.getDepartment())) {
						List<Holiday> tempholidays = holidayPersons.get(i)
								.getHolidays();
						tempholidays.add(holiday);
						holidayPersons.get(i).setHolidays(tempholidays);
						if (holiday.getType().equals("年假")) {
							holidayPersons.get(i).setNianjia(
									holidayPersons.get(i).getNianjia()
											+ holiday.getHours());
						} else if (holiday.getType().equals("事假")) {
							holidayPersons.get(i).setShijia(
									holidayPersons.get(i).getShijia()
											+ holiday.getHours());
						} else if (holiday.getType().equals("病假")) {
							holidayPersons.get(i).setBingjia(
									holidayPersons.get(i).getBingjia()
											+ holiday.getHours());
						} else if (holiday.getType().equals("调休")) {
							holidayPersons.get(i).setTiaoxiu(
									holidayPersons.get(i).getTiaoxiu()
											+ holiday.getHours());
						} else {
							holidayPersons.get(i).setQita(
									holidayPersons.get(i).getQita()
											+ holiday.getHours());
						}
						j = i;
						break;
					}
				}
				if (!holidayPersons.get(j).getName().equals(holiday.getName())
						|| !holidayPersons.get(j).getDepartment()
								.equals(holiday.getDepartment())) {
					HolidayPerson tempholidayPerson = new HolidayPerson();
					tempholidayPerson.setName(holiday.getName());
					tempholidayPerson.setDepartment(holiday.getDepartment());
					List<Holiday> tempholidays = new ArrayList<Holiday>();
					tempholidays.add(holiday);
					tempholidayPerson.setHolidays(tempholidays);
					if (holiday.getType().equals("年假")) {
						tempholidayPerson.setNianjia(holiday.getHours());
					} else if (holiday.getType().equals("事假")) {
						tempholidayPerson.setShijia(holiday.getHours());
					} else if (holiday.getType().equals("病假")) {
						tempholidayPerson.setBingjia(holiday.getHours());
					} else if (holiday.getType().equals("调休")) {
						tempholidayPerson.setTiaoxiu(holiday.getHours());
					} else {
						tempholidayPerson.setQita(holiday.getHours());
					}
					holidayPersons.add(tempholidayPerson);
				}
			} else {
				HolidayPerson holidayPerson = new HolidayPerson();
				holidayPerson.setName(holiday.getName());
				holidayPerson.setDepartment(holiday.getDepartment());
				List<Holiday> tempholidays = new ArrayList<Holiday>();
				tempholidays.add(holiday);
				holidayPerson.setHolidays(tempholidays);
				if (holiday.getType().equals("年假")) {
					holidayPerson.setNianjia(holiday.getHours());
				} else if (holiday.getType().equals("事假")) {
					holidayPerson.setShijia(holiday.getHours());
				} else if (holiday.getType().equals("病假")) {
					holidayPerson.setBingjia(holiday.getHours());
				} else if (holiday.getType().equals("调休")) {
					holidayPerson.setTiaoxiu(holiday.getHours());
				} else {
					holidayPerson.setQita(holiday.getHours());
				}
				holidayPersons.add(holidayPerson);
			}

		}

		SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
		Date start = sdf.parse(date);

		for (HolidayPerson holidayPerson : holidayPersons) {
			String fileName1 = holidayPerson.getDepartment() + "-"
					+ holidayPerson.getName();
			for (Holiday holiday : holidayPerson.getHolidays()) {
				String fileName = holiday.getDepartment() + "-"
						+ holiday.getName();
				Date end = sdf.parse(holiday.getDate());
				long diff = end.getTime() - start.getTime();
				int rowNum = (int) (diff / 1000 / 60 / 60 / 24 + 5);

				Tools.writerString(outPath, fileName, outPath, fileName, XLS,
						rowNum, 10, holiday.getStart(), true);
				Tools.writerString(outPath, fileName, outPath, fileName, XLS,
						rowNum, 11, holiday.getEnd(), true);
				Tools.writerDouble(outPath, fileName, outPath, fileName, XLS,
						rowNum, 12, holiday.getHours(), true);
				Tools.writerString(outPath, fileName, outPath, fileName, XLS,
						rowNum, 13, holiday.getType(), true);

			}
			Tools.writerDouble(outPath, fileName1, outPath, fileName1, XLS, 39,
					1, holidayPerson.getNianjia(), false);
			Tools.writerDouble(outPath, fileName1, outPath, fileName1, XLS, 39,
					4, holidayPerson.getBingjia(), false);
			Tools.writerDouble(outPath, fileName1, outPath, fileName1, XLS, 39,
					7, holidayPerson.getShijia(), false);
			Tools.writerDouble(outPath, fileName1, outPath, fileName1, XLS, 39,
					10, holidayPerson.getTiaoxiu(), false);
			Tools.writerDouble(outPath, fileName1, outPath, fileName1, XLS, 39,
					13, holidayPerson.getQita(), false);

		}
	}

	public static void sumPBs() throws IOException {
		List<List<Object>> names = Tools.readAll(inPath, name, XLSX);
		List<Person> persons = new ArrayList<Person>();
		for (int i = 1; i < names.size(); i++) {
			Person person = new Person();
			person.setName((String) names.get(i).get(1));
			person.setDepartment((String) names.get(i).get(0));
			persons.add(person);
		}
		for (Person person : persons) {
			String fileName = person.getDepartment() + "-" + person.getName();
			List<List<Object>> details = Tools.readAll(outPath, fileName, XLS);
			int zhong = 0;
			int ye = 0;
			for (int i = 5; i < 36; i++) {
				if (((String) details.get(i).get(2)).contains("中班")) {
					zhong++;
				} else if (((String) details.get(i).get(2)).contains("晚班")) {
					ye++;
				}
			}
			Tools.writerString(outPath, fileName, outPath, fileName, XLS, 38,
					1, "" + zhong, false);
			Tools.writerString(outPath, fileName, outPath, fileName, XLS, 38,
					4, "" + ye, false);
		}
	}

	public static void setOuts() throws IOException, ParseException {
		List<List<Object>> outs = Tools.readAll(inPath, out, XLS);

		List<Out> outList = new ArrayList<Out>();
		for (int i = 3; i < outs.size(); i++) {
			Out out = new Out();
			out.setDate(((String) outs.get(i).get(4)).substring(5));
			out.setDepartment((String) outs.get(i).get(2));
			out.setName((String) outs.get(i).get(1));

			String start = ((String) outs.get(i).get(5)).substring(11);
			String end = ((String) outs.get(i).get(6)).substring(11);

			SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");

			Date start1 = sdf.parse(start);
			Date end1 = sdf.parse(end);
			int time = 0;
			if (start1.before(sdf.parse("09:01"))) {
				time += 1;
			}
			if (end1.after(sdf.parse("17:29"))) {
				time += 2;
			}
			out.setTime(time);

			outList.add(out);

		}
		for (Out out : outList) {
			String fileName = out.getDepartment() + "-" + out.getName();
			Workbook wb = Tools.open(outPath, fileName, XLS);
			Sheet sheet = wb.getSheetAt(0);
			String date = out.getDate();
			for (int i = 5; i <= sheet.getLastRowNum(); i++) {
				if (sheet.getRow(i).getCell(0).getStringCellValue()
						.equals(date)) {
					switch (out.getTime()) {
					case 1:
						if (sheet.getRow(i).getCell(3).getStringCellValue()
								.contains("缺勤")) {
							Tools.writerString(outPath, fileName, outPath,
									fileName, XLS, i, 3, "外出", true);
						}
						break;
					case 2:
						if (sheet.getRow(i).getCell(4).getStringCellValue()
								.contains("缺勤")) {
							Tools.writerString(outPath, fileName, outPath,
									fileName, XLS, i, 4, "外出", true);
						}
						break;
					case 3:
						if (sheet.getRow(i).getCell(3).getStringCellValue()
								.contains("缺勤")) {
							Tools.writerString(outPath, fileName, outPath,
									fileName, XLS, i, 3, "外出", true);
						}
						if (sheet.getRow(i).getCell(4).getStringCellValue()
								.contains("缺勤")) {
							Tools.writerString(outPath, fileName, outPath,
									fileName, XLS, i, 4, "外出", true);
						}
						break;
					default:
						break;
					}
				}
			}
		}

	}

	public static void setAdds() throws IOException, ParseException {
		List<List<Object>> add1s = Tools.readAll(inPath, add1, XLS);
		List<List<Object>> add2s = Tools.readAll(inPath, add2, XLS);
		List<List<Object>> add2ses = Tools.readAll(inPath, add2se, XLS);

		List<Add> adds = new ArrayList<Add>();

		for (int i = 3; i < add2s.size(); i++) {
			Add add = new Add();
			add.setDate(((String) add2s.get(i).get(4)).substring(5, 10));
			add.setDepartment((String) add2s.get(i).get(2));
			add.setName((String) add2s.get(i).get(1));
			add.setStart(((String) add2s.get(i).get(4)).substring(5));
			add.setEnd(((String) add2s.get(i).get(5)).substring(5));
			add.setSite((String) add2s.get(i).get(7));
			add.setApply("是");
			double hours = (Double) add2s.get(i).get(6) * 24;
			BigDecimal b = new BigDecimal(hours);
			add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());

			adds.add(add);
		}

		for (int i = 3; i < add2ses.size(); i++) {
			for (int j = 0; j < adds.size(); j++) {
				if (adds.get(j).getName()
						.equals((String) add2ses.get(i).get(1))
						&& adds.get(j).getDepartment()
								.equals((String) add2ses.get(i).get(2))
						&& adds.get(j)
								.getStart()
								.equals(((String) add2ses.get(i).get(4))
										.substring(5))
						&& adds.get(j)
								.getEnd()
								.equals(((String) add2ses.get(i).get(5))
										.substring(5))) {
					adds.get(j).setApply("否");
					break;
				}
			}
		}

		for (int i = 3; i < add1s.size(); i++) {
			Add add = new Add();
			add.setDate(((String) add1s.get(i).get(3)).substring(5, 10));
			add.setDepartment((String) add1s.get(i).get(2));
			add.setName((String) add1s.get(i).get(1));
			add.setStart(((String) add1s.get(i).get(3)).substring(5));
			add.setEnd(((String) add1s.get(i).get(4)).substring(5));
			add.setSite((String) add1s.get(i).get(7));
			add.setApply("是");
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
			Date start = sdf.parse((String) add1s.get(i).get(3));
			Date end = sdf.parse((String) add1s.get(i).get(4));
			long diff = end.getTime() - start.getTime();
			double hours = (double) diff / 60 / 60 / 1000;
			BigDecimal b = new BigDecimal(hours);
			add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());

			adds.add(add);
		}

		for (Add add : adds) {
			String fileName = add.getDepartment() + "-" + add.getName();
			Workbook wb = Tools.open(outPath, fileName, XLS);
			Sheet sheet = wb.getSheetAt(0);
			String date = add.getDate();

			for (int i = 5; i <= sheet.getLastRowNum(); i++) {
				if (sheet.getRow(i).getCell(0).getStringCellValue()
						.equals(date)) {
					if (sheet.getRow(i).getCell(6).getStringCellValue() == "") {
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i, 5, add.getSite(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i, 6, add.getStart(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i, 7, add.getEnd(), true);
						Tools.writerDouble(outPath, fileName, outPath,
								fileName, XLS, i, 8, add.getHours(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i, 9, add.getApply(), true);
					} else {
						Tools.shift(outPath, fileName, outPath, fileName, XLS,
								i + 1);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i + 1, 5, add.getSite(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i + 1, 6, add.getStart(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i + 1, 7, add.getEnd(), true);
						Tools.writerDouble(outPath, fileName, outPath,
								fileName, XLS, i + 1, 8, add.getHours(), true);
						Tools.writerString(outPath, fileName, outPath,
								fileName, XLS, i + 1, 9, add.getApply(), true);
					}
				}
			}
		}
	}
}

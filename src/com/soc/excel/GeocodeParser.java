package com.soc.excel;

import java.util.ArrayList;

/**
 * 스트링 문자열 주소를 반환하는 Class 객체 생성후 Parser 메서드를 이용하여 문자열을 잘라내 준다.
 * 
 * @author zuneho
 * 
 */
public class GeocodeParser{
	public GeocodeParser() {
	}

	/**
	 * 문자열 사이의 구분자를 이용하여 문자열을 자르고 주소체계에 맞게 arrayList로 반환하여 준다.
	 * 
	 * @param msg
	 *            주소를 찾아낼 문자열
	 * @param splitString
	 *            주소 사이에 잘라낼 때 사용 할 구분자
	 * @param splitTarget
	 *            주소 레벨 (0~4까지 설정) 0전체 1 시군구 2 읍면동 3 리 레벨로 주소를 구한다.
	 * @return String 으로 구성된 ArrayList며 각 배열에 들어간 값은 다음과 같다 0=시도 1=시군구 2=읍면동 3=리
	 *         4=산 5=본번 6=부번 7=나머지 주소
	 */
	public ArrayList<String> Parser(String msg, String splitString, int splitLevel) {
		String text = msg;// .replace(" ", "~");
		String DO = " ";
		String SI = " ";
		String DONG = " ";
		String RE = " ";
		String SAN = " ";
		String BON = " ";
		String BU = " ";
		String ADDR = " ";
		ArrayList<String> result = new ArrayList<String>();
		String[] address = text.split(splitString);
		// SplitLevel 에 따라 부분적인 주소만을 추려낸다. 0전체 1 시군구 2 읍면동 3 리
		if (splitLevel != 0) {
			DO = "_";
		}
		if (splitLevel != 0 && splitLevel != 1) {
			SI = "_";
		}
		if (splitLevel != 0 && splitLevel != 1 && splitLevel != 2) {
			DONG = "_";
		}
		for ( int i = 0;i < address.length;i++) {
			String splitText = address[i];
			if (DO.equals(" ") && !(DO.equals("_")) && (splitText.lastIndexOf("시") != -1 || splitText.lastIndexOf("도") != -1)) {
				DO = splitText;
				if (DO.lastIndexOf("광주") != -1) {
					DO = "광주광역시";
				} else if (DO.lastIndexOf("대구") != -1) {
					DO = "대구광역시";
				} else if (DO.lastIndexOf("대전") != -1) {
					DO = "대전광역시";
				} else if (DO.lastIndexOf("부산") != -1) {
					DO = "부산광역시";
				} else if (DO.lastIndexOf("서울") != -1) {
					DO = "서울특별시";
				} else if (DO.lastIndexOf("세종") != -1) {
					DO = "세종특별자치시";
				} else if (DO.lastIndexOf("울산") != -1) {
					DO = "울산광역시";
				} else if (DO.lastIndexOf("인천") != -1) {
					DO = "인천광역시";
				} else if (DO.lastIndexOf("제주") != -1) {
					DO = "제주특별자치도";
				}
			}
			if (!(DO.equals(" ")) && SI.equals(" ") && (splitText.lastIndexOf("시") != -1 || splitText.lastIndexOf("군") != -1 || splitText.lastIndexOf("구") != -1)) {
				if (DO.lastIndexOf("_") != -1 && splitText.lastIndexOf("시") != -1) {
				} else if (DO.lastIndexOf("시") != -1 && splitText.lastIndexOf("시") != -1) {
				} else if (DO.lastIndexOf("도") != -1 && splitText.lastIndexOf("시") != -1) {
					SI = splitText;
				} else {
					SI = splitText;
				}
			} else if (!(DO.equals(" ")) && SI.lastIndexOf("시") >= 0 && (DO.lastIndexOf("도") >= 0 || DO.lastIndexOf("_") >= 0)) {
				if (!(SI.equals("_")) && splitText.lastIndexOf("구") != -1 && splitText.lastIndexOf("구") == splitText.length() - 1) {
					SI = SI + splitText;
				}
			}
			if (!(SI.equals(" ")) && !(DONG.equals("_")) && (splitText.lastIndexOf("읍") != -1 || splitText.lastIndexOf("면") != -1 || splitText.lastIndexOf("동") != -1)) {
				DONG = splitText;
			}
			if (!(DONG.equals(" ")) && !(DONG.equals("_")) && splitText.lastIndexOf("가") != -1 && splitText.substring(splitText.length() - 1, splitText.length()).equals("가")) {
				DONG = DONG + splitText;
			}
			if (!(DONG.equals(" ")) && splitText.lastIndexOf("리") != -1) {
				RE = splitText;
			}
			// 기존 주소
			if (!(SAN.equals(" ")) && splitText.lastIndexOf("산") != -1 && splitText.length() < 3) {
				System.out.println(splitText);
				System.out.println(splitText.lastIndexOf("산"));
				System.out.println(splitText.length());
				if (splitText.length() > 1) {
					SAN = splitText.substring(0, splitText.lastIndexOf("산"));
				} else {
					SAN = "2";
				}
			} else if (SAN.equals(" ")) {
				SAN = "1";
			}
			if (splitText.lastIndexOf("-") != -1) {
				BON = splitText.substring(0, splitText.lastIndexOf("-"));
				BU = splitText.substring(splitText.lastIndexOf("-") + 1, splitText.length());
			}
			if (i == address.length - 1 && splitText.lastIndexOf("-") == -1) {
				ADDR = splitText;
			}
		}
		if (RE.equals(" ")) {
			RE = "_";
		}
		result.add(DO);
		result.add(SI);
		result.add(DONG);
		result.add(RE);
		result.add(SAN);
		result.add(BON);
		result.add(BU);
		result.add(ADDR);
		return result;
	}
}

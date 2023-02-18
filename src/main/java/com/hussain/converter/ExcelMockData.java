package com.hussain.converter;

import com.hussain.response.ExcelVo;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.List;


@Component
public class ExcelMockData {

	private List<ExcelVo> excelData;

	public List<ExcelVo> getExcelData() {
		if (CollectionUtils.isEmpty(excelData)) {
			populateExcelData();
		}
		return excelData;
	}

	private void populateExcelData() {
		excelData = new ArrayList<>();
		Class<ExcelVo> classZ = (Class<ExcelVo>) ExcelVo.class;
		Method[] methods = classZ.getMethods();
		for (int i = 0; i < 20000; i++) {
			ExcelVo model = new ExcelVo();
			for (Method method : methods) {
				String methodName = method.getName();
				if (methodName.startsWith("set")) {
					try {
						method.invoke(model, getRandomString());
					} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
						e.printStackTrace();
					}
				}
			}
			excelData.add(model);
		}
	}

	private String getRandomString() {
		SecureRandom random = new SecureRandom();
		return new BigInteger(130, random).toString(32);
	}

}

package com.xiaobei.util;


/**
 * ExcelUtils的一个行映射器
 * 
 */
public interface ExcelUtilsRowMapper {

	Object[] rowMapping(Object record) throws Exception;
}

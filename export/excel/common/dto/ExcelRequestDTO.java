/*
 * Copyright (c) 2019 
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 */
package export.excel.common.dto;

import java.io.Serializable;
import java.rmi.Remote;
import java.util.EnumMap;
import java.util.List;

import javax.security.auth.callback.Callback;

import export.excel.common.constants.ExcelConstant;
import export.excel.common.interfaces.ExcelDataSource;
import export.excel.common.model.ExcelCustomAdd;

/**
 * 
 * The {@code ExcelRequestDTO} class represents data to send to be generated for
 * the excel report ExcelRequestDTOs are constant and cannot be changed after
 * they are created. ExcelRequestDTOs cannot be shared and are thread safe.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.model.ExcelCustomAdd
 * @see export.excel.common.interfaces.ExcelDataSource
 * @see java.util.EnumMap
 * @since JDK1.5
 */

public final class ExcelRequestDTO implements Callback, Serializable, Remote {

	/** use serialVersionUID from JDK 1.0.2 for interoperability */

	private static final long serialVersionUID = 4434944939748189239L;

	/** The value is used for row headings storage for excel */
	private final List<String> rowHeadingLS;

	/** The value is used for default header footer for excel */
	private final EnumMap<ExcelConstant, Object> defaultMap;

	/** The value is used for custom implementation for excel */
	private final ExcelCustomAdd excelCustomAdd;

	/** The value receives any object that implements <ExcelDataSource> */
	private final List<? extends ExcelDataSource> excelDataSource;

	/** The value is used for custom index definition of columns */
	private final int[] dataIndex;
	

	/**
	 * Initializes a newly created {@code ExcelRequestDTO} object so that it
	 * represents {@code defaultMap} {@code rowHeadingLS} {@code excelDataSource}
	 * {@code dataIndex}. The use of this constructor is to create an object with
	 * default excel implementation and custom index definition.
	 */
	public ExcelRequestDTO(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			List<? extends ExcelDataSource> excelDataSource, int[] dataIndex) {
		this.rowHeadingLS = rowHeadingLS;
		this.defaultMap = defaultMap;
		this.excelCustomAdd = null;
		this.excelDataSource = excelDataSource;
		this.dataIndex = dataIndex;

	}

	/**
	 * Initializes a newly created {@code ExcelRequestDTO} object so that it
	 * represents {@code defaultMap} {@value rowHeadingLS} {@code excelDataSource}
	 * {@code dataIndex}. The use of this constructor is to create an object with
	 * default excel and index definition implementation.
	 */
	public ExcelRequestDTO(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			List<? extends ExcelDataSource> excelDataSource) {

		this.rowHeadingLS = rowHeadingLS;
		this.defaultMap = defaultMap;
		this.excelCustomAdd = null;
		this.excelDataSource = excelDataSource;
		this.dataIndex = null;
	}

	/**
	 * Initializes a newly created {@code ExcelRequestDTO} object so that it
	 * represents {@code defaultMap} {@value rowHeadingLS} {@value excelDataSource}
	 * {@value dataIndex}. The use of this constructor is to create an object with
	 * custom excel implementation and custom index position definition
	 */
	public ExcelRequestDTO(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			List<? extends ExcelDataSource> excelDataSource, ExcelCustomAdd excelCustomAdd, int[] dataIndex) {
		super();
		this.rowHeadingLS = rowHeadingLS;
		this.defaultMap = defaultMap;
		this.excelCustomAdd = excelCustomAdd;
		this.excelDataSource = excelDataSource;
		this.dataIndex = dataIndex;
	}

	/**
	 * Initializes a newly created {@code ExcelRequestDTO} object so that it
	 * represents {@value defaultMap} {@value rowHeadingLS} {@value excelDataSource}
	 * {@value dataIndex} . The use of this constructor is to create an object with
	 * custom excel implementation and and without custom index position definition
	 */
	public ExcelRequestDTO(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			List<? extends ExcelDataSource> excelDataSource, ExcelCustomAdd excelCustomAdd) {
		super();
		this.rowHeadingLS = rowHeadingLS;
		this.defaultMap = defaultMap;
		this.excelCustomAdd = excelCustomAdd;
		this.excelDataSource = excelDataSource;
		dataIndex = null;
	}

	/**
	 * Returns a {@value rowHeadingLS}.
	 */
	public List<String> getRowHeadingLS() {
		return rowHeadingLS;
	}

	/**
	 * Returns a {@value defaultMap}.
	 */
	public EnumMap<ExcelConstant, Object> getDefaultMap() {
		return defaultMap;
	}

	/**
	 * Returns a {@value excelDataSource}.
	 */
	public List<? extends ExcelDataSource> getExcelDataSource() {
		return excelDataSource;
	}

	/**
	 * Returns a {@value dataIndex}.
	 */
	public int[] getDataIndex() {
		return dataIndex;
	}

	/**
	 * Returns a {@value excelCustomAdd}.
	 */
	public ExcelCustomAdd getExcelCustomAdd() {
		return excelCustomAdd;
	}


}

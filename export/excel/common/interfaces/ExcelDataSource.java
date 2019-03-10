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
package export.excel.common.interfaces;

import java.io.Serializable;
import java.rmi.Remote;
import javax.security.auth.callback.Callback;
/**
 * 
 * The {@code ExcelDataSourceO} is a marker interface. The interface should
 * be implemented by every class which has to be printed on Excel.
 * It will help to generalize every bean as a data source.
 * 
 * @author Ruba Mushtaq
 * @since JDK1.5
 */


public interface ExcelDataSource extends Serializable,Callback,Remote{

}

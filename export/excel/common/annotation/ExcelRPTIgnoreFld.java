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
 */
package export.excel.common.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
/**
 * 
 * The {@code ExcelRPTIgnoreFld} placed on fields which should
 * be ignored while reading data in excel on runtime.
 * 
 * <blockquote><pre>
 *    Class ABC{
 *    @ExcelRPTIgnoreFld
 *     private String s1; // wont be shown in excel
 *     
 *     private String s2;
 *    
 *    }
 * </pre></blockquote>
 * 
 * @author Ruba Mushtaq
 * @since JDK1.5
 */

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelRPTIgnoreFld {

}

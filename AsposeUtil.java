package org.springblade.common.utils;

import com.aspose.cells.Workbook;
import com.aspose.pdf.DocSaveOptions;
import com.aspose.pdf.ExcelSaveOptions;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springblade.core.log.exception.ServiceException;
import org.springblade.modules.resource.controller.UploadController;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;

/**
 * @author zhs
 * @Description
 * @createTime 2021/5/21 0021 13:40
 */
@Component
@Slf4j
public class AsposeUtil {




	public  boolean getWordLicense() {
		boolean result = false;
		try {
			InputStream is = UploadController.class.getClassLoader().getResourceAsStream("license.xml");
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	public  boolean getCellsLicense() {
		boolean result = false;
		try {
			InputStream is = UploadController.class.getClassLoader().getResourceAsStream("license.xml");
			com.aspose.cells.License aposeLic = new com.aspose.cells.License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}


	/**
	 * word转pdf
	 * @param is 文件输入流
	 * @return 文件
	 */
	public String word2Pdf(InputStream is, String fileName) throws Exception {
		getWordLicense();
		Document doc = new Document(is);

		String tempPath = System.getProperty("user.dir").concat("/upload/");
		doc.save(tempPath.concat(fileName), SaveFormat.PDF);

		is.close();
		return tempPath.concat(fileName);
	}

	/**
	 * excel转pdf
	 * @param is
	 * @param fileName
	 * @return
	 * @throws Exception
	 */
	public String excel2Pdf(InputStream is, String fileName) throws Exception {
		getCellsLicense();
		Workbook workbook2 = new Workbook(is);
		String tempPath = System.getProperty("user.dir").concat("/upload/");
		workbook2.save(tempPath.concat(fileName), com.aspose.cells.SaveFormat.PDF);
		is.close();
		return tempPath.concat(fileName);
	}



	public String file2Pdf(InputStream is, String fileName) throws Exception {
		String extensionName = StringUtils.substringAfterLast(fileName, ".");
		fileName = StringUtils.substringBeforeLast(fileName, ".").concat(".pdf");
		if (Arrays.asList("xls", "xlsx").contains(extensionName.toLowerCase())) {
			//excel 转pdf
			return excel2Pdf(is, fileName);
		}else if (Arrays.asList("doc", "docx").contains(extensionName.toLowerCase())){
			// word 转pdf
			return word2Pdf(is, fileName);
		}else {
			return null;
		}
	}
}


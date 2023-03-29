import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.TempFile;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class PPTXBinder {
    XMLSlideShow origin;
    XMLSlideShow dest;

    public PPTXBinder(String url) {
        // get File Object from resourceUrl
        File file = new File(url);

        try (FileInputStream is = new FileInputStream(file)) {
            origin = new XMLSlideShow(is);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private void _bindText(XSLFSlide slide, XSLFTextShape originShape) {
        XSLFTextShape destShape = slide.createTextBox();

        destShape.setAnchor(originShape.getAnchor());
        destShape.setFillColor(originShape.getFillColor());
        destShape.setLineColor(originShape.getLineColor());
        destShape.setLineWidth(originShape.getLineWidth());
        destShape.setLineDash(originShape.getLineDash());
        destShape.setLineCap(originShape.getLineCap());
        destShape.setLineCompound(originShape.getLineCompound());

        for (XSLFTextParagraph paragraph : originShape.getTextParagraphs()) {
            XSLFTextParagraph destParagraph = destShape.addNewTextParagraph();
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                String text = run.getRawText();
                destParagraph.addNewTextRun().setText(text);
            }
        }
    }

    private void _bindImage(XSLFSlide slide, XSLFPictureShape originShape) {
        // XSLFPictureShape destShape = slide.createPicture(originShape.getPictureData());

        // destShape.setAnchor(originShape.getAnchor());
    }

    private void _bindTable(XSLFSlide slide, XSLFTable originShape) {
        XSLFTable destShape = slide.createTable();

        destShape.setAnchor(originShape.getAnchor());

        for (XSLFTableRow originRow : originShape.getRows()) {
            XSLFTableRow destRow = destShape.addRow();

            destRow.setHeight(originRow.getHeight());

            for (XSLFTableCell originCell : originRow.getCells()) {
                XSLFTextShape destCell = destRow.addCell();

                destCell.setAnchor(originCell.getAnchor());
                destCell.setFillColor(originCell.getFillColor());
                destCell.setLineColor(originCell.getLineColor());
                destCell.setLineWidth(originCell.getLineWidth());
                destCell.setLineDash(originCell.getLineDash());
                destCell.setLineCap(originCell.getLineCap());
                destCell.setLineCompound(originCell.getLineCompound());

                for (XSLFTextParagraph paragraph : originCell.getTextParagraphs()) {
                    XSLFTextParagraph destParagraph = destCell.addNewTextParagraph();
                    for (XSLFTextRun run : paragraph.getTextRuns()) {
                        String text = run.getRawText();
                        destParagraph.addNewTextRun().setText(text);
                    }
                }
            }
        }
    }

    public File bind(String data) throws ParseException {
        JSONParser parser = new JSONParser();
        parser.parse(data);

        dest = new XMLSlideShow();
        for (XSLFSlide originSlide : origin.getSlides()) {
            XSLFSlide slide = dest.createSlide();

            for (XSLFShape originShape : originSlide.getShapes()) {
                if (originShape instanceof XSLFTextShape) {
                    _bindText(slide, (XSLFTextShape) originShape);
                } else if (originShape instanceof XSLFPictureShape) {
                    _bindImage(slide, (XSLFPictureShape) originShape);
                } else if (originShape instanceof XSLFTable) {
                    _bindTable(slide, (XSLFTable) originShape);
                }
            }
        }

        File temp;
        try {
            temp = TempFile.createTempFile("binded", ".pptx");
            FileOutputStream fos = new FileOutputStream(temp);
            dest.write(fos);
            fos.flush();
            fos.close();
            origin.close();

            return temp;
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return null;
    }

	public static void main(String[] args) {
		try {
			PPTXBinder binder = new PPTXBinder("bldProposal.pptx");
			File file = binder.bind(
					"[{\"buildingId\":\"BLD00038613\",\"buildingName\":\"내광빌딩\",\"buildingImage\":\"https://image-proxy-dev.rsquareon.com/x/photos/2017/9/19/9/38613/AvzDBy6rMK/original/700-16.jpg\",\"address\":\"서울특별시 강남구 역삼동 700-16\",\"roadNameAddress\":\"언주로94길 7\",\"subwayStationInformation\":\"선릉역 (수인분당선) 도보 7분\",\"totalAreaM2\":1281.9,\"totalArea\":387.77,\"floorCount\":5,\"basementCount\":1,\"mainPurpose\":\"제2종근린생활시설\",\"buildingDirection\":\"남\",\"rawCompletedConstructDate\":\"19831104\",\"remodelingYear\":\"2002\",\"exclusiveRate\":78,\"standardLeasableAreaM2\":231.404,\"standardLeasableArea\":70,\"standardNetLeasableAreaM2\":297.52,\"standardNetLeasableArea\":90,\"elevatorTotalCount\":1,\"totalParkingCount\":null,\"freeParkingDetail\":\"1층당 1대\",\"paidParkingDetail\":null,\"products\":[{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000},{\"buildingId\":\"BLD00038613\",\"rawUnitCfCd\":\"01\",\"rawDgUseYn\":\"N\",\"rawDgName\":null,\"rawFlrNum\":4,\"rawUnitName\":\"전체\",\"leasableAreaM2\":297.52,\"netLeasableAreaM2\":231.404,\"rawVacancy\":\"01\",\"rawMvinPslbYm\":null,\"deposit\":200000000,\"rent\":6000000,\"maintenanceFee\":1000000}],\"totalLeasableAreaM2\":3900.822,\"totalNetLeasableAreaM2\":3061.152,\"rawCoolerState\":\"02\",\"rawHeaterState\":\"02\",\"photos\":[[null]],\"userName\":\"정재희\",\"latitude\":37.5045594,\"longitude\":127.0426748,\"public\":1,\"freight\":0,\"outline\":{\"outlineId\":\"11680-12422\",\"buildingName\":null,\"insideParkingLotCount\":0,\"outsideParkingLotCount\":5,\"insideMechanicalParkingLotCount\":0,\"outsideMechanicalParkingLotCount\":0,\"summaryParkingCount\":0},\"rawAllDayOpen\":\"02\",\"floors\":null,\"floorTotalLeasableArea\":0,\"floorTotalLeasableAreaM2\":0,\"floorTotalNetLeasableArea\":0,\"floorTotalNetLeasableAreaM2\":0,\"floorNOC\":0}]");

			// copy file to another location
			FileInputStream in = new FileInputStream(file);
			FileOutputStream out = new FileOutputStream(new File("~/Downloads/test.pptx"));
			byte[] buf = new byte[1024];
			int len;
			while ((len = in.read(buf)) > 0) {
				out.write(buf, 0, len);
			}
			in.close();
			out.close();
			System.out.println("File copied.");
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

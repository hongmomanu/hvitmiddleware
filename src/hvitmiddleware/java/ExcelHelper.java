package hvitmiddleware.java;

import jxl.Workbook;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: jack
 * Date: 13-9-9
 * Time: 下午1:45
 * To change this template use File | Settings | File Templates.
 */
public class ExcelHelper {
    private static final Logger log = Logger.getLogger(ExcelHelper.class);

    public static Map<String, Object> writeExcel(String fileName, String header_arr, String rowdata, String sum,
                                                 String title, int headerheight,int headercols,boolean isall,String url,String extraParams,String rowname,String pager) {
        WritableWorkbook wwb = null;
        Map<String, Object> map = new HashMap<String, Object>();
        int sumrow_index = 0;
        map.put("isok", false);
        try {
            //首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
            File file = new File(fileName);
            file.getParentFile().mkdirs();
            file.createNewFile();
            wwb = Workbook.createWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        JSONObject sum_item = JSONObject.fromObject(sum);
        if (wwb != null) {
            //创建一个可写入的工作表
            //Workbook的createSheet方法有两个参数，第一个是工作表的名称，第二个是工作表在工作薄中的位置
            WritableSheet ws = wwb.createSheet("sheet1", 0);
            JSONArray headers = JSONArray.fromObject(header_arr);
            JSONArray pagers= JSONArray.fromObject(pager);
            JSONArray rowdatas=new JSONArray();
            if(isall){
               JSONObject urlparam=JSONObject.fromObject(extraParams);
                String urlparamsstr="";
                for (Object param_name : urlparam.names()) {
                    String value=urlparam.get(param_name).toString();
                    if(value.equals("null"))continue;
                    try {
                        urlparamsstr+=param_name.toString()+"="+ URLEncoder.encode(value, "UTF-8")+"&";
                    } catch (UnsupportedEncodingException e) {
                        e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
                    }

                }
                //urlparamsstr+="limit=0&start=-1";

                rowdatas=JSONObject.fromObject(UrlConnectHelper.sendPost(url,urlparamsstr)).getJSONArray(rowname);
                
                /**
                Map<String,Double> sum_new=new HashMap<String, Double>();
                for (Object sum_name : sum_item.names()) {
                    sum_new.put(sum_name.toString(),0.0);

                }

                for (int row_index = 0; row_index < rowdatas.size(); row_index++) {

                    for (Object sum_name : sum_item.names()) {
                                String val=JSONObject.fromObject(rowdatas.get(row_index)).getString(sum_name.toString());
                                if(val.equals("null")){
                                val="0";
                                					}
                        sum_new.put(sum_name.toString(),sum_new.get(sum_name.toString())+
                                Double.parseDouble(val));

                    }
                    //String division=JSONObject.fromObject(rowdatas.get(row_index)).getString("division");
                    //String sql="select parentid,divisionname from divisions where divisionpath MATCH '"+division+"'";
                    //ComonDao cd=new ComonDao();
                    //Map<String,Object> item=cd.getSigleObj(sql);
                    //int parentid=Integer.parseInt(item.get("parentid").toString());
                    //String divisionname=item.get("divisionname").toString();
                    //ArrayList<String>result=new ArrayList<String>();
                    //result.add(divisionname);
                    //result=getDivisionTreeBypath(parentid,"divisions",result);
                    //JSONObject row=JSONObject.fromObject(rowdatas.get(row_index));
                    /**for(int i=result.size()-1;i>=0;i--){
                       if(i==(result.size()-1)){
                           row.put("city",result.get(i));
                       }
                       if(i==(result.size()-2)){
                           row.put("county",result.get(i));
                       }
                       if(i==(result.size()-3)){
                           row.put("town",result.get(i));
                       }
                       if(i==(result.size()-4)){
                           row.put("village",result.get(i));
                       }

                    }
                    //rowdatas.set(row_index, row);

                }
                sum_item=JSONObject.fromObject(sum_new);**/

            }else{
                rowdatas = JSONArray.fromObject(rowdata);
            }

            try {
                WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                        15,
                        WritableFont.BOLD,
                        false,
                        UnderlineStyle.NO_UNDERLINE);
                ws.mergeCells(0, 0, headercols - 1, 0);
                Label labelTitle = new Label(0, 0, title);
                WritableCellFormat cellFormat = new WritableCellFormat();
                cellFormat.setAlignment(jxl.format.Alignment.CENTRE);
                cellFormat.setFont(font);
                labelTitle.setCellFormat(cellFormat);
                ws.addCell(labelTitle);
            } catch (WriteException e) {
                e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
            }
            WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                    10,
                    WritableFont.BOLD,
                    false,
                    UnderlineStyle.NO_UNDERLINE);
            makemultiheader(ws, headers, 0, rowdatas, sum_item, sumrow_index,headerheight,pagers);

            try {
                //从内存中写入文件中
                wwb.write();
                //关闭资源，释放内存
                wwb.close();
                map.put("isok", true);
                map.put("path", fileName);

            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        }
        return map;
    }

    
    public static void makemultiheader(WritableSheet ws, JSONArray headers,
                                       int colindex, JSONArray rowdatas, JSONObject sum_item, int sumrow_index,int headerheight,JSONArray pagers) {
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                10,
                WritableFont.BOLD,
                false,
                UnderlineStyle.NO_UNDERLINE);
        for (int j = colindex; j < headers.size()+colindex; j++) {
            if(headers.getJSONObject(j-colindex).getJSONArray("columns").size()>0){
                makemultiheader(ws,headers.getJSONObject(j-colindex).getJSONArray("columns"),j,rowdatas,sum_item,sumrow_index,headerheight,pagers);
            }

            try {
            			//System.out.println(j);
            			//System.out.println(headers);
            			//System.out.println(headers.size());
                String col_name = headers.getJSONObject(j-colindex).getString("value");
                //添加表头
                WritableCellFormat cellFormat = new WritableCellFormat();
                cellFormat.setAlignment(jxl.format.Alignment.CENTRE);
                cellFormat.setFont(font);

                JSONArray cols=headers.getJSONObject(j-colindex).getJSONArray("col");
                JSONArray rows=headers.getJSONObject(j-colindex).getJSONArray("row");
                ws.mergeCells(cols.getInt(0), rows.getInt(0), cols.getInt(1), rows.getInt(1));

                Label labelC = new Label(cols.getInt(0), rows.getInt(0), headers.getJSONObject(j-colindex).getString("name"));
                labelC.setCellFormat(cellFormat);
                ws.addCell(labelC);


                //添加行数据
                for (int row_index = 0; row_index < rowdatas.size(); row_index++) {
                    WritableCellFormat cellRowFormat = new WritableCellFormat();
                    cellRowFormat.setAlignment(jxl.format.Alignment.CENTRE);
                    sumrow_index = row_index + headerheight+1;
                    Label labelRowC = null;

                    if (col_name.equals("index")) {
                        labelRowC = new Label(cols.getInt(0), sumrow_index, String.valueOf(row_index + 1));

                    } else {

                            if(rowdatas.getJSONObject(row_index).has(col_name)){
                                labelRowC = new Label(cols.getInt(0), sumrow_index,rowdatas.getJSONObject(row_index).
                                        getString(col_name).equals("null")?"":rowdatas.getJSONObject(row_index).
                                        getString(col_name));
                            }
                            else continue;

                    }
                    labelRowC.setCellFormat(cellRowFormat);
                    ws.addCell(labelRowC);

                }

                //添加合计数据
                sumrow_index++;
                Label labelSumC = null;
                WritableCellFormat cellRowFormat = new WritableCellFormat();
                cellRowFormat.setAlignment(jxl.format.Alignment.CENTRE);
											String tempval="";
                if (j == 0) {
                    tempval="合计:";
                    labelSumC = new Label(j, sumrow_index, "合计");
                    labelSumC.setCellFormat(cellRowFormat);
                    ws.addCell(labelSumC);

                } 
                    for (Object sum_name : sum_item.names()) {
																//System.out.println(col_name);
																//System.out.println(sum_name);
                        if (col_name.equals(sum_name.toString())) {
                            labelSumC = new Label(cols.getInt(0), sumrow_index, tempval+sum_item.get(sum_name).toString());
                            
                           
                        }else{
                        labelSumC = new Label(cols.getInt(0), sumrow_index, tempval);
                        
                        }
                        labelSumC.setCellFormat(cellRowFormat);
                            ws.addCell(labelSumC);
                         break;
                    }
                


                //添加表单数据
                if (j == 0) {
                    sumrow_index++;
                    
                    
                    
                     for (int pager_index = 0; pager_index < pagers.size(); pager_index++) {
                     JSONArray pager_cols=pagers.getJSONObject(pager_index).getJSONArray("col");
                     JSONArray pager_rows=pagers.getJSONObject(pager_index).getJSONArray("row");
                     ws.mergeCells(pager_cols.getInt(0), sumrow_index+pager_rows.getInt(0), pager_cols.getInt(1), sumrow_index+pager_rows.getInt(1));
                     
                     Label labelLast_head_C = new Label(pager_cols.getInt(0), sumrow_index+pager_rows.getInt(0), pagers.getJSONObject(pager_index).getString("name")+pagers.getJSONObject(pager_index).getString("value"));
                    WritableCellFormat cellLastRowHeaderFormat = new WritableCellFormat();
                    cellLastRowHeaderFormat.setAlignment(Alignment.LEFT);
                    cellLastRowHeaderFormat.setFont(font);
                    labelLast_head_C.setCellFormat(cellLastRowHeaderFormat);
                    ws.addCell(labelLast_head_C);
                     //System.out.println(pagers.getJSONObject(pager_index));
                     
                     
                     }
                    
                    /**
                    ws.mergeCells(0, sumrow_index, headers.size() / 2 - 1, sumrow_index);
                    ws.mergeCells(headers.size() / 2, sumrow_index, headers.size() - 1, sumrow_index);
                    Label labelLast_head_C = new Label(0, sumrow_index, "填表人:          分管领导:");
                    WritableCellFormat cellLastRowHeaderFormat = new WritableCellFormat();
                    cellLastRowHeaderFormat.setAlignment(Alignment.LEFT);
                    cellLastRowHeaderFormat.setFont(font);
                    labelLast_head_C.setCellFormat(cellLastRowHeaderFormat);
                    ws.addCell(labelLast_head_C);

                    String date_str = StringHelper.getTimeStrFormat("yyyy-MM-dd");
                    Label labelLast_tail_C = new Label(headers.size() / 2, sumrow_index, "填表日期: " + date_str);
                    WritableCellFormat cellLastRowTailFormat = new WritableCellFormat();
                    cellLastRowTailFormat.setAlignment(Alignment.RIGHT);
                    cellLastRowTailFormat.setFont(font);
                    labelLast_tail_C.setCellFormat(cellLastRowTailFormat);
                    ws.addCell(labelLast_tail_C);**/
                    
                    


                }


            } catch (RowsExceededException e) {
                log.debug(e.getMessage());
            } catch (WriteException e) {
                log.debug(e.getMessage());
            }

        }


    }


}

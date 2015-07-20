import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.StringWriter;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import javax.xml.transform.*;
import javax.xml.transform.dom.*;
import javax.xml.transform.stream.*;

/**
 * 
 */

/**
 * @author drashtti
 * 
 * this tool takes a spreadsheet and
 * converts it into a .mwt format
 * that can be used as a dictionary 
 * compatible to the WhatIzIt. 
 *
 */
public class Converter {

	/**
	 * @param args
	 */
	
	
	private Document doc; 
	private Element root;
	
	public void readFile( String filelocation){
		try {
			Workbook wbk1 = WorkbookFactory
					.create(new FileInputStream(
							filelocation));
			Sheet s1 = wbk1.getSheetAt(0);
			Workbook wbk2 = WorkbookFactory.create(new FileInputStream("/home/drashtti/Desktop/ontologies/Diabetes-Onto/synhgnc.xls"));
			
			Sheet s2 = wbk2.getSheetAt(0);
			
			Sheet s3 = wbk2.getSheetAt(1);
			System.out.println("The sheet 3 has been initiated and it is " + s3);
			int rownum = s1.getLastRowNum();
			for (int n = 1; n <= rownum; n++) {
				//first row is column headers hence skipped 
				Row r1 = s1.getRow(n);
				System.out.println("Row r1 is initiated and is " + r1);
				Row r2 = s2.getRow(n);
				System.out.println("Row r2 is initiated and is " + r2);
				Row r3 = s3.getRow(n);
				System.out.println("Row r3 is initiated and is " + r3);
				//get the ID 
				Cell id = r1.getCell(0);
				String hgncid = id.toString();
				
				//get the symbol 
				Cell symbol = r1.getCell(1);
				String symb = symbol.toString();
				createXML(hgncid,symb);
				
				//get the name 
				Cell label = r1.getCell(2);
				String name = label.toString();
				createXML(hgncid,name);
				
				//get the symbol syns - sheet 1 and till row m 
				if(r2 != null){
				for (int x = 0; x <=12 ; x++){
					if(r2.getCell(x) != null){
					Cell symsyn = r2.getCell(x);
					String value = symsyn.toString();
					createXML(hgncid,value);
					}
					}
				}
				
				//get the name syns - sheet 2 and till row k 
				if(r3 != null){
				for (int y = 0; y<=10 ; y++){
					if(r3.getCell(y)!= null){
						Cell namesyn = r3.getCell(y);
						String value = namesyn.toString();
						createXML(hgncid,value);
					}
					}
				}
				
				
			}
			
			
		} catch (InvalidFormatException e) {
			System.out.println("Please check the format of the input");
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			System.out.println("The input file was not found");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("something went wrong with IO");
			e.printStackTrace();
		}
		
	}
	
	
	public void createHeaderXML(){
		
	        try {
	            
	            //Creating an empty XML Document

	            //We need a Document
	            DocumentBuilderFactory dbfac = DocumentBuilderFactory.newInstance();
	            DocumentBuilder docBuilder = dbfac.newDocumentBuilder();
	            doc = docBuilder.newDocument();
	            //create the root element and add it to the document
	            root = doc.createElement("mwt");
	            doc.appendChild(root);
	            //create the template element <template><z:hgnc ids='%1'>%0</z:hgnc></template>
	            Element template = doc.createElement("template");
	            root.appendChild(template);
	            Element zhgnc = doc.createElement("z:hgnc");
	            zhgnc.setAttribute("ids", "%1");
	            template.appendChild(zhgnc);
	            Text text = doc.createTextNode("%0");
	            zhgnc.appendChild(text);

	        }catch (Exception e) {
	            System.out.println(e);
	        }
	}
	
	
	public void createXML(String hgncid, String value){
		
        //create child element, add an attribute, and add to root
        Element child = doc.createElement("t");
        child.setAttribute("p1", hgncid);
        root.appendChild(child);
        //add a text element to the child
        Text text = doc.createTextNode(value);
        child.appendChild(text);
	}
	
	public void createFootXML(){
		/**<template>%0</template>
		<r>&lt;z:[^&gt;]*&gt;(.*&lt;/z)!:[^&gt;]*></r>
		<r>&lt;protname[^&gt;]*&gt;(.*&lt;/protname)![^&gt;]*></r>
		<!-- next one to skip anything looking like an html tag -->
		<r>&lt;/?[A-Za-z_0-9\-]+(&gt;|[\r\n\t ][^&gt;]+)</r>*/
		
		  Element template = doc.createElement("template");
          root.appendChild(template);
          Text text = doc.createTextNode("%0");
          template.appendChild(text);
          
          Element r1 = doc.createElement("r");
          root.appendChild(r1);
          Text text1 = doc.createTextNode("&lt;z:[^&gt;]*&gt;(.*&lt;/z)!:[^&gt;]*>");
          r1.appendChild(text1);
          
          Element r2 = doc.createElement("r");
          root.appendChild(r2);
          Text text2 = doc.createTextNode("&lt;protname[^&gt;]*&gt;(.*&lt;/protname)![^&gt;]*>");
          r2.appendChild(text2);

          Element r3 = doc.createElement("r");
          root.appendChild(r3);
          Text text3 = doc.createTextNode("&lt;/?[A-Za-z_0-9\\-]+(&gt;|[\r\n\t ][^&gt;]+)");
          r3.appendChild(text3);
		
	}
	
	public void saveXML (){
		
		// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer;
				try {
					transformer = transformerFactory.newTransformer();
					transformer.setOutputProperty(OutputKeys.INDENT, "yes");
					transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new File("/home/drashtti/Desktop/ontologies/hgncGene.mwt"));
		 
				// Output to console for testing
				// StreamResult result = new StreamResult(System.out);
		 
				transformer.transform(source, result);
		 
				System.out.println("File saved!");
				}
				catch (TransformerConfigurationException e) {
					System.out.println(e);
					e.printStackTrace();
				} catch (TransformerException e) {
					System.out.println(e);
					e.printStackTrace();
				}
		
	}
	
	public static void main(String[] args) {
		Converter convert = new Converter();
		convert.createHeaderXML();
		convert.readFile("/home/drashtti/Desktop/ontologies/Diabetes-Onto/HGNC.xls");
		convert.createFootXML();
		convert.saveXML();
	}

}

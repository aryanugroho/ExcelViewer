package com.fndcz.excel;

import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * Servlet implementation class ExcelViewerServlet
 */
@WebServlet("/view")
public class ExcelViewerServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * Default constructor. 
     */
    public ExcelViewerServlet() {
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
		StringBuilder html = new StringBuilder();
		
		ToHtml toHtml = ToHtml.create("c:/status.xlsx", html);

		List<String> names = toHtml.getSheetNames();
		
		response.setHeader("Content-Type", "text/html;charset=utf-8");
		response.setCharacterEncoding("UTF-8");
		response.getWriter().println("<!DOCTYPE html>"
				+ "<html>"
				+ "<head>"
				+ "<meta charset=\"utf-8\">"
				+ "</head>"
				+ "<body>"
				+ "<form action=\"./view\" method=\"POST\">"
				+ "<select SIZE=\"10\" name=\"sheet_index\">");
		int index = 0;
		for(String name : names) {
			response.getWriter().println("<option value=\"" + index + "\">" + name + "</option>");
			index++;
		}

		response.getWriter().println(
				"</select>"
				+ "<br/><input type='password' placeholder='password' name='password'/>"
				+ "<br/><input type='submit' value='view'/>"
				+ "</form>"
				+ "</body>"
				+ "</html>");
		
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String indexStr = request.getParameter("sheet_index");
		String password = request.getParameter("password");
		
		if("xxxxx".equals(password)) {
			
			int sheetIndex = 0; 
			try {
				sheetIndex = Integer.valueOf(indexStr);
			} catch(Exception e) {
				
			}
			
			StringBuilder html = new StringBuilder();
			ToHtml toHtml = ToHtml.create("c:/status.xlsx", html);
			toHtml.setCompleteHTML(true);
			toHtml.printPage(sheetIndex);
			
			response.setHeader("Content-Type", "text/html;charset=utf-8");
			response.setCharacterEncoding("UTF-8");
			response.getWriter().write(html.toString());
		}
	}

}

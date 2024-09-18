package main

import (
	"bytes"
	"fmt"
	"gh0ffice/lib"
	"html"
	"os"
	"regexp"
	"strings"

	// "github.com/ledongthuc/pdf"
	"github.com/moipa-cn/pptx"
	"github.com/nguyenthenguyen/docx"
	"github.com/thedatashed/xlsxreader"
	"seehuhn.de/go/pdf"
)

var TAG_RE = regexp.MustCompile(`(<[^>]*>)+`)
var PARA_RE = regexp.MustCompile(`(</[a-z]:p>)+`)

func main() {
	docx2txt("data/1.docx")
	pptx2txt("data/2.pptx")
	xlsx2txt("data/3.xlsx")
	// pdf2txt("data/4.pdf")
	doc2txt("data/1.doc")
	ppt2txt("data/2.ppt")
	xls2txt("data/3.xls")
	pdf2txt("data/4_1.pdf")
}

func removeStrangeChars(input string) string {
	// Define the regex pattern for allowed characters
	re := regexp.MustCompile("[�\x13\x0b]+")
	// Replace all disallowed characters with an empty string
	return re.ReplaceAllString(input, " ")
}

func docx2txt(filename string) ([]string, error) {
	data_docx, err := docx.ReadDocxFile(filename) // Read data from docx file
	if err != nil {
		return []string{}, err
	}
	text_docx := data_docx.Editable().GetContent() // Get whole docx data as XML formated text
	paras_docx := PARA_RE.Split(text_docx, -1)     // Split the docx in paragraphs
	for i := range paras_docx {                    // For each paragraph
		paragraph := TAG_RE.ReplaceAllString(paras_docx[i], "") // Remove all the tags to extract the content
		paragraph = html.UnescapeString(paragraph)              // replace all the html entities (e.g. &amp)
		paras_docx[i] = paragraph
		// fmt.Println(i, removeStrangeChars(paragraph))
	}
	data_docx.Close()
	return paras_docx, nil
}

func pptx2txt(filename string) ([]string, error) {
	data_pptx, err := pptx.ReadPowerPoint(filename) // Read data from pptx file
	if err != nil {
		return []string{}, err
	}

	slides_pptx := data_pptx.GetSlidesContent() // Get pptx slides data as an array of XML formated text
	var paras_pptx []string
	for i := range slides_pptx {
		slide_paras_pptx := PARA_RE.Split(slides_pptx[i], -1) // Split the docx in paragraphs
		for j := range slide_paras_pptx {                     // For each paragraph
			paragraph := TAG_RE.ReplaceAllString(slide_paras_pptx[j], "") // Remove all the tags to extract the content
			paragraph = html.UnescapeString(paragraph)                    // replace all the html entities (e.g. &amp)
			if len(paragraph) > 0 {
				paras_pptx = append(paras_pptx, paragraph) // Save all paragraphs as ONE array
				// fmt.Println(i, j, removeStrangeChars(paragraph))
			}
		}
	}
	return paras_pptx, nil
}

func xlsx2txt(filename string) ([]string, error) {
	data_xlsx, err := xlsxreader.OpenFile(filename) // Read data from xlsx file
	if err != nil {
		return []string{}, err
	}
	var rows_xlsx []string
	for _, sheet := range data_xlsx.Sheets { // For each sheet of the file
		for row := range data_xlsx.ReadRows(sheet) { // For each row of the sheet
			text_row := ""
			for i, col := range row.Cells { // Concatenate cells of the row with tab separator
				if i > 0 {
					text_row = fmt.Sprintf("%s\t%s", text_row, col.Value)
				} else {
					text_row = fmt.Sprintf("%s%s", text_row, col.Value)
				}
			}
			rows_xlsx = append(rows_xlsx, text_row)
			// fmt.Println(removeStrangeChars(text_row))
		}
	}
	data_xlsx.Close()
	return rows_xlsx, nil
}

func pdf2txt(filename string) ([]string, error) {
	ropt := &pdf.ReaderOptions{
		ReadPassword:  func(ID []byte, try int) string { return "" },
		ErrorHandling: pdf.ErrorHandlingReport,
	}
	r, err := pdf.Open(filename, ropt)
	if err != nil {
		return []string{}, err
	}
	r.Close()
	return []string{}, nil
}

// func pdf2txt(filename string) ([]string, error) {
// 	file_pdf, data_pdf, err := pdf.Open(filename) // Read data from pdf file
// 	if err != nil {
// 		return []string{}, err
// 	}
// 	var pages_pdf []string
// 	for i := 1; i <= data_pdf.NumPage(); i++ { // For each page of the file
// 		data_page := data_pdf.Page(i) // Get page data
// 		if data_page.V.IsNull() {     // If page is empty, move to the next iteration
// 			continue
// 		}
// 		rows_page, err := data_page.GetTextByRow() // Get text by rows in the page
// 		if err != nil {                            // If an error occurred, move to the next iteration
// 			return []string{}, err
// 		}
// 		text_page := ""
// 		for j, row_page := range rows_page { // For each row
// 			text_row := ""
// 			for _, word_row := range row_page.Content { // Concatenate words of the row
// 				text_row = fmt.Sprintf("%s%s", text_row, word_row.S)
// 			}
// 			// Concatenate rows of the page
// 			if j > 0 {
// 				text_page = fmt.Sprintf("%s\n%s", text_page, text_row)
// 			} else {
// 				text_page = fmt.Sprintf("%s%s", text_page, text_row)
// 			}
// 		}
// 		pages_pdf = append(pages_pdf, removeStrangeChars(text_page))
// 		fmt.Println(removeStrangeChars(text_page))
// 	}
// 	file_pdf.Close()
// 	return pages_pdf, nil
// }

func doc2txt(filename string) ([]string, error) {
	file_doc, _ := os.Open(filename)        // Open doc file
	data_doc, err := lib.DOC2Text(file_doc) // Read data from a doc file
	if err != nil {
		return []string{}, err
	}

	actual := data_doc.(*bytes.Buffer) // Buffer for hold line text of doc file
	var paras_doc []string
	for aline, err := actual.ReadString('\r'); err == nil; aline, err = actual.ReadString('\r') { // Get text by line
		aline = strings.Trim(aline, " \n\r")
		if aline != "" {
			paras_doc = append(paras_doc, removeStrangeChars(aline))
			// fmt.Println(removeStrangeChars(aline))
		}
	}
	file_doc.Close()
	return paras_doc, nil
}

func ppt2txt(filename string) ([]string, error) {
	file_ppt, err := os.Open(filename) // Open ppt file
	if err != nil {
		return []string{}, err
	}

	text_ppt, err := lib.ExtractText(file_ppt) // Read text from a ppt file
	if err != nil {
		return []string{}, err
	}

	var paras_ppt []string
	for _, aline := range strings.Split(text_ppt, "\r") { // Seperate text as lines
		aline = strings.Trim(aline, " \n\r")
		if aline != "" {
			paras_ppt = append(paras_ppt, removeStrangeChars(aline))
			// fmt.Println(removeStrangeChars(aline))
		}
	}
	file_ppt.Close()
	return paras_ppt, nil
}

func xls2txt(filename string) ([]string, error) {
	file_xls, err := os.Open(filename) // Open xls file
	if err != nil {
		return []string{}, err
	}

	rows_xls, err := lib.XLS2Text(file_xls) // Convert xls data to an array of rows (include all sheets)
	if err != nil {
		return []string{}, err
	}
	// for _, row := range rows_xls {
	// 	fmt.Println("🍺", row)
	// }
	file_xls.Close()
	return rows_xls, nil
}

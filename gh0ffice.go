/*
 Licensed to the Apache Software Foundation (ASF) under one
 or more contributor license agreements.  See the NOTICE file
 distributed with this work for additional information
 regarding copyright ownership.  The ASF licenses this file
 to you under the Apache License, Version 2.0 (the
 "License"); you may not use this file except in compliance
 with the License.  You may obtain a copy of the License at
   http://www.apache.org/licenses/LICENSE-2.0
 Unless required by applicable law or agreed to in writing,
 software distributed under the License is distributed on an
 "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 KIND, either express or implied.  See the License for the
 specific language governing permissions and limitations
 under the License.
*/

package gh0ffice

import (
	"bytes"
	"errors"
	"fmt"
	"html"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"
	"syscall"
	"time"

	"github.com/WhityGhost/gh0ffice/lib"
	"github.com/WhityGhost/gh0ffice/lib/metagoffice"
	"github.com/WhityGhost/gh0ffice/lib/pdf"

	"github.com/charmbracelet/log"
	"github.com/moipa-cn/pptx"
	"github.com/nguyenthenguyen/docx"
	"github.com/thedatashed/xlsxreader"
)

const ISO string = "2006-01-02T15:04:05"

var TAG_RE = regexp.MustCompile(`(<[^>]*>)+`)
var PARA_RE = regexp.MustCompile(`(</[a-z]:p>)+`)
var DEBUG bool = false

type Document struct {
	path           string
	RePath         string    `json:"path"`
	Filename       string    `json:"filename"`
	Title          string    `json:"title"`
	Subject        string    `json:"subject"`
	Creator        string    `json:"creator"`
	Keywords       string    `json:"keywords"`
	Description    string    `json:"description"`
	Lastmodifiedby string    `json:"lastModifiedBy"`
	Revision       string    `json:"revision"`
	Category       string    `json:"category"`
	Content        string    `json:"content"`
	Modifytime     time.Time `json:"modified"`
	Createtime     time.Time `json:"created"`
	Accesstime     time.Time `json:"accessed"`
	Size           int       `json:"size"`
}

type DocReader func(string) (string, error)

func SetDebug(dbg bool) {
	DEBUG = dbg
}

// Make a struct of documentation involves content and metadata, file information
func InspectDocument(pathname string, target_abpath string) (*Document, error) {
	abPath, err := filepath.Abs(pathname)
	if err != nil {
		return nil, err
	}
	rePath := abPath
	rePath = strings.TrimPrefix(rePath, target_abpath)
	filename := path.Base(pathname)
	data := Document{path: pathname, RePath: rePath, Title: filename}
	extension := path.Ext(pathname)
	_, err = insertFileInfoData(&data)
	if err != nil {
		return &data, err
	}
	switch extension {
	case ".docx":
		_, e := insertMetaData(&data)
		if e != nil && DEBUG {
			log.Warnf("⚠️ %s", e.Error())
		}
		_, err = insertContentData(&data, docx2txt)
	case ".pptx":
		_, e := insertMetaData(&data)
		if e != nil && DEBUG {
			log.Warnf("⚠️ %s", e.Error())
		}
		_, err = insertContentData(&data, pptx2txt)
	case ".xlsx":
		_, e := insertMetaData(&data)
		if e != nil && DEBUG {
			log.Warnf("⚠️ %s", e.Error())
		}
		_, err = insertContentData(&data, xlsx2txt)
	case ".pdf":
		_, err = insertContentData(&data, pdf2txt)
	case ".doc":
		_, err = insertContentData(&data, doc2txt)
	case ".ppt":
		_, err = insertContentData(&data, ppt2txt)
	case ".xls":
		_, err = insertContentData(&data, xls2txt)
	}
	if err != nil {
		return &data, err
	}
	if DEBUG {
		log.Infof("✔️ successfully read content of file: %s", data.Filename)
		printFileInfoData(&data)
	}
	return &data, nil
}

// Read the meta data of office files (only *.docx, *.xlsx, *.pptx) and insert into the interface
func insertMetaData(data *Document) (bool, error) {
	file, err := os.Open(data.path)
	if err != nil {
		return false, err
	}
	defer file.Close()
	meta, err := metagoffice.GetContent(file)
	if err != nil {
		return false, errors.New("failed to get office meta data")
	}
	if meta.Title != "" {
		data.Title = meta.Title
	}
	data.Subject = meta.Subject
	data.Creator = meta.Creator
	data.Keywords = meta.Keywords
	data.Description = meta.Description
	data.Lastmodifiedby = meta.LastModifiedBy
	data.Revision = meta.Revision
	data.Category = meta.Category
	data.Content = meta.Category
	return true, nil
}

// Read the content of office files and insert into the interface
func insertContentData(data *Document, reader DocReader) (bool, error) {
	content, err := reader(data.path)
	if err != nil {
		return false, err
	}
	data.Content = content
	return true, nil
}

// Read the file information of any files and insert into the interface
func insertFileInfoData(data *Document) (bool, error) {
	fileinfo, err := os.Stat(data.path)
	if err != nil {
		return false, err
	}
	// if runtime.GOOS == "windows" {
	stat := fileinfo.Sys().(*syscall.Win32FileAttributeData)
	data.Filename = fileinfo.Name()
	data.Title = data.Filename
	data.Createtime = time.Unix(0, stat.LastAccessTime.Nanoseconds())
	data.Modifytime = time.Unix(0, stat.CreationTime.Nanoseconds())
	data.Accesstime = time.Unix(0, stat.LastWriteTime.Nanoseconds())
	data.Size = int(fileinfo.Size())
	// } else {
	// 	aTime := fileinfo.Sys().(*syscall.Stat_t).Atim
	// 	cTime := fileinfo.Sys().(*syscall.Stat_t).Ctim
	// 	mTime := fileinfo.Sys().(*syscall.Stat_t).Mtim
	// 	data = int(aTime.Nsec)
	// 	data = int(cTime.Nsec)
	// 	data = int(mTime.Nsec)
	// }
	return true, nil
}

// Print the file information (except for content) for debugging
func printFileInfoData(data *Document) {
	if data.Filename != "" {
		log.Infof("📄 filename: %s", data.Filename)
	}
	if data.Title != "" {
		log.Infof("📄 title: %s", data.Title)
	}
	if data.Subject != "" {
		log.Infof("📄 subject: %s", data.Subject)
	}
	if data.Creator != "" {
		log.Infof("👤 creator: %s", data.Creator)
	}
	if data.Keywords != "" {
		log.Infof("🗝️ keywords: %s", data.Keywords)
	}
	if data.Description != "" {
		log.Infof("📄 description: %s", data.Description)
	}
	if data.Lastmodifiedby != "" {
		log.Infof("👤 lastmodifiedby: %s", data.Lastmodifiedby)
	}
	if data.Revision != "" {
		log.Infof("📄 revision: %s", data.Revision)
	}
	if data.Category != "" {
		log.Infof("📄 category: %s", data.Category)
	}
	if !data.Modifytime.IsZero() {
		log.Infof("📆 modifytime (ISO): %s", data.Modifytime.Format(ISO))
	}
	if !data.Createtime.IsZero() {
		log.Infof("📆 createtime (ISO): %s", data.Createtime.Format(ISO))
	}
	if !data.Accesstime.IsZero() {
		log.Infof("📆 accesstime (ISO): %s", data.Accesstime.Format(ISO))
	}
	// if data.content != "" {
	//	log.Infof("📄 content: %s", data.content))
	// }
}

func removeStrangeChars(input string) string {
	// Define the regex pattern for allowed characters
	re := regexp.MustCompile("[�\x13\x0b]+")
	// Replace all disallowed characters with an empty string
	return re.ReplaceAllString(input, " ")
}

func docx2txt(filename string) (string, error) {
	data_docx, err := docx.ReadDocxFile(filename) // Read data from docx file
	if err != nil {
		return "", err
	}
	defer data_docx.Close()
	text_docx := data_docx.Editable().GetContent()        // Get whole docx data as XML formated text
	text_docx = PARA_RE.ReplaceAllString(text_docx, "\n") // Replace the end of paragraphs (</w:p) with /n
	text_docx = TAG_RE.ReplaceAllString(text_docx, "")    // Remove all the tags to extract the content
	text_docx = html.UnescapeString(text_docx)            // Replace all the html entities (e.g. &amp)

	// fmt.Println(text_docx)
	return text_docx, nil
}

func pptx2txt(filename string) (string, error) {
	data_pptx, err := pptx.ReadPowerPoint(filename) // Read data from pptx file
	if err != nil {
		return "", err
	}

	data_pptx.DeletePassWord()
	slides_pptx := data_pptx.GetSlidesContent() // Get pptx slides data as an array of XML formated text
	var text_pptx string
	for i := range slides_pptx {
		slide_text_pptx := PARA_RE.ReplaceAllString(slides_pptx[i], "\n") // Replace the end of paragraphs (</w:p) with /n
		slide_text_pptx = TAG_RE.ReplaceAllString(slide_text_pptx, "")    // Remove all the tags to extract the content
		slide_text_pptx = html.UnescapeString(slide_text_pptx)            // Replace all the html entities (e.g. &amp)
		if slide_text_pptx != "" {                                        // Save all slides as ONE string
			if text_pptx != "" {
				text_pptx = fmt.Sprintf("%s\n%s", text_pptx, slide_text_pptx)
			} else {
				text_pptx = fmt.Sprintf("%s%s", text_pptx, slide_text_pptx)
			}
		}
	}
	// fmt.Println(text_pptx)
	return text_pptx, nil
}

func xlsx2txt(filename string) (string, error) {
	data_xlsx, err := xlsxreader.OpenFile(filename) // Read data from xlsx file
	if err != nil {
		return "", err
	}
	defer data_xlsx.Close()

	var rows_xlsx string
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
			if rows_xlsx != "" { // Save all rows as ONE string
				rows_xlsx = fmt.Sprintf("%s\n%s", rows_xlsx, text_row)
			} else {
				rows_xlsx = fmt.Sprintf("%s%s", rows_xlsx, text_row)
			}
		}
	}
	// fmt.Println(rows_xlsx)
	return rows_xlsx, nil
}

func pdf2txt(filename string) (string, error) { // BUG: Cannot get text from specific (or really malformed?) pages
	file_pdf, data_pdf, err := pdf.Open(filename) // Read data from pdf file
	if err != nil {
		return "", err
	}
	defer file_pdf.Close()

	var buff_pdf bytes.Buffer
	bytes_pdf, err := data_pdf.GetPlainText() // Get text of entire pdf file
	if err != nil {
		return "", err
	}

	buff_pdf.ReadFrom(bytes_pdf)
	text_pdf := buff_pdf.String()
	// fmt.Println(text_pdf)
	return text_pdf, nil
}

func doc2txt(filename string) (string, error) {
	file_doc, _ := os.Open(filename)        // Open doc file
	data_doc, err := lib.DOC2Text(file_doc) // Read data from a doc file
	if err != nil {
		return "", err
	}
	defer file_doc.Close()

	actual := data_doc.(*bytes.Buffer) // Buffer for hold line text of doc file
	text_doc := ""
	for aline, err := actual.ReadString('\r'); err == nil; aline, err = actual.ReadString('\r') { // Get text by line
		aline = strings.Trim(aline, " \n\r")
		if aline != "" {
			if text_doc != "" {
				text_doc = fmt.Sprintf("%s\n%s", text_doc, removeStrangeChars(aline))
			} else {
				text_doc = fmt.Sprintf("%s%s", text_doc, removeStrangeChars(aline))
			}
		}
	}
	text_doc = removeStrangeChars(text_doc)
	// fmt.Println(text_doc)
	return text_doc, nil
}

func ppt2txt(filename string) (string, error) {
	file_ppt, err := os.Open(filename) // Open ppt file
	if err != nil {
		return "", err
	}
	defer file_ppt.Close()

	text_ppt, err := lib.ExtractText(file_ppt) // Read text from a ppt file
	if err != nil {
		return "", err
	}
	text_ppt = removeStrangeChars(text_ppt)
	// fmt.Println(text_ppt)
	return text_ppt, nil
}

func xls2txt(filename string) (string, error) {
	file_xls, err := os.Open(filename) // Open xls file
	if err != nil {
		return "", err
	}
	defer file_xls.Close()

	text_xls, err := lib.XLS2Text(file_xls) // Convert xls data to an array of rows (include all sheets)
	if err != nil {
		return "", err
	}
	text_xls = removeStrangeChars(text_xls)
	// fmt.Println(text_xls)
	return text_xls, nil
}

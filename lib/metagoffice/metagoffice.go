package metagoffice

import (
	"archive/zip"
	"bufio"
	"encoding/xml"
	"errors"
	"os"
)

// XMLContent contains the fields of te file core.xml
type XMLContent struct {
	Title          string `xml:"title"`
	Subject        string `xml:"subject"`
	Creator        string `xml:"creator"`
	Keywords       string `xml:"keywords"`
	Description    string `xml:"description"`
	LastModifiedBy string `xml:"lastModifiedBy"`
	Revision       string `xml:"revision"`
	Created        string `xml:"created"`
	Modified       string `xml:"modified"`
	Category       string `xml:"category"`
}

// GetContent function
func GetContent(document *os.File) (fields XMLContent, err error) {
	// Attempt to read the document file directly as a zip file.
	z, err := zip.OpenReader(document.Name())
	if err != nil {
		return fields, errors.New("failed to open the file as zip")
	}
	defer z.Close()

	var xmlFile string
	for _, file := range z.File {
		if file.Name == "docProps/core.xml" {
			rc, err := file.Open()
			if err != nil {
				return fields, errors.New("failed to open docProps/core.xml")
			}
			defer rc.Close()

			scanner := bufio.NewScanner(rc)
			for scanner.Scan() {
				xmlFile += scanner.Text()
			}
			if err := scanner.Err(); err != nil {
				return fields, errors.New("failed to read from docProps/core.xml")
			}
			break // Exit loop after finding and reading core.xml
		}
	}

	// Unmarshal the collected XML content into the XMLContent struct
	if err := xml.Unmarshal([]byte(xmlFile), &fields); err != nil {
		return fields, errors.New("failed to Unmarshal")
	}

	return fields, nil
}

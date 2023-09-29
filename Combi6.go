package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"regexp"
	"strconv"
	"strings"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

func main() {
	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatalf("Error al leer el directorio: %v", err)
	}

	categories := []string{"fcc", "fex", "fc"}
	workbooks := make(map[string]*excelize.File)

	for _, category := range categories {
		workbooks[category] = excelize.NewFile()
		firstFile := true
		sheetName := "Sheet1"

		for _, file := range files {
			filename := file.Name()
			baseName := strings.TrimSuffix(filename, ".xlsx")

			if strings.Contains(baseName, category) && !strings.Contains(baseName, category+"c") && strings.HasSuffix(filename, ".xlsx") {
				log.Printf("Añadiendo archivo %s a la categoría '%s'\n", filename, category)
				processFile(filename, workbooks[category], sheetName, &firstFile)
			}
		}

		if !firstFile {
			outputFilename := fmt.Sprintf("documento combinado %s %s.xlsx", category, time.Now().Format("2006-01-02_15-04-05"))
			if err := workbooks[category].SaveAs(outputFilename); err != nil {
				log.Printf("Error al guardar el archivo combinado %s: %v\n", outputFilename, err)
			} else {
				log.Printf("Archivo combinado creado: %s\n", outputFilename)
			}
		}
	}
}

func processFile(filename string, combinedFile *excelize.File, sheetName string, firstFile *bool) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatalf("Error al abrir el archivo %s: %v", filename, err)
	}

	sourceSheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sourceSheetName)
	if err != nil {
		log.Fatalf("Error al obtener filas del archivo %s: %v", filename, err)
	}

	existingRows, _ := combinedFile.GetRows(sheetName)
	startRowIndex := len(existingRows) + 1

	if *firstFile {
		combinedFile.SetCellValue(sheetName, "A1", "ID")
		combinedFile.SetSheetRow(sheetName, "B1", &rows[0])
		*firstFile = false
	}
	// Skip headers for subsequent files
	rows = rows[1:]

	for _, row := range rows {
		for i := range row {
			row[i] = cleanHTMLContent(row[i])
		}
		lastIdentifier := getLastIdentifier(combinedFile, sheetName) + 1
		newRow := append([]string{strconv.Itoa(lastIdentifier)}, row...)
		axis := fmt.Sprintf("A%d", startRowIndex)
		combinedFile.SetSheetRow(sheetName, axis, &newRow)
		startRowIndex++
	}
}

func cleanHTMLContent(content string) string {
	re := regexp.MustCompile(`</?[a-z][\s\S]*?>`)
	return re.ReplaceAllString(content, "")
}

func getLastIdentifier(combinedFile *excelize.File, sheetName string) int {
	lastRow, _ := combinedFile.GetRows(sheetName)
	lastIdentifier := 0
	if len(lastRow) > 0 && len(lastRow[len(lastRow)-1]) > 0 {
		lastIdentifier, _ = strconv.Atoi(lastRow[len(lastRow)-1][0])
	}
	return lastIdentifier
}

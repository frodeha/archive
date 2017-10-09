package Archive

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize"
	"golang.org/x/text/unicode/norm"
)

var (
	dropboxURL = "https://www.dropbox.com/home/Note arkivet"
	columns    = []string{"Navn", "Komponist", "Arrangør", "Partitur", "Treblås", "Messing", "Slagverk"}
)

// ArchiveKey is a key
type ArchiveKey struct {
	file     *excelize.File
	fileName string
	name     string
	rows     []ArchiveRow
}

func NewArchiveKey(fileName string) *ArchiveKey {
	return &ArchiveKey{fileName: fileName, name: "Note arkivet", rows: make([]ArchiveRow, 0)}
}

func (a *ArchiveKey) Rows() []ArchiveRow { return a.rows }

type ArchiveRow struct {
	index           int
	Name            string
	comparableField string
	Composer        string
	Arranger        string
	HasScore        bool
	HasWoodwind     bool
	HasBrass        bool
	HasPercussion   bool
}

func (a *ArchiveKey) AddRow(name, composer, arranger string, hasScore, hasWoodwind, hasBrass, hasPercussion bool) {

	normalizedName := string(norm.NFC.Bytes([]byte(name)))

	row := ArchiveRow{
		Name:            normalizedName,
		comparableField: strings.ToLower(normalizedName),
		Arranger:        arranger,
		Composer:        composer,
		HasScore:        hasScore,
		HasWoodwind:     hasWoodwind,
		HasBrass:        hasBrass,
		HasPercussion:   hasPercussion,
	}

	for i, v := range a.rows {

		if row.comparableField == v.comparableField {
			return
		}

		if row.comparableField < v.comparableField {

			a.rows = append(a.rows[:i],
				append(
					[]ArchiveRow{row}, a.rows[i:]...,
				)...,
			)

			return
		}
	}

	a.rows = append(a.rows, row)

}

func (a *ArchiveKey) UpdateRow(r ArchiveRow) error {

	for i, v := range a.rows {
		if v.comparableField == r.comparableField {
			row := &a.rows[i]
			row.Composer = r.Composer
			row.Arranger = r.Arranger
			row.HasScore = r.HasScore
			row.HasWoodwind = r.HasWoodwind
			row.HasBrass = r.HasBrass
			row.HasPercussion = r.HasPercussion
			return nil
		}
	}

	return fmt.Errorf("Found no row with the same name as the provided row")
}

func (a *ArchiveKey) GetRow(name string) (ArchiveRow, error) {
	comparableField := strings.ToLower(string(norm.NFC.Bytes([]byte(name))))

	for _, v := range a.rows {
		if v.comparableField == comparableField {
			return v, nil
		}
	}

	return ArchiveRow{}, fmt.Errorf("Found no row with the same name as the provided row")
}

func (a *ArchiveKey) DeleteRow(name string) error {
	comparableField := strings.ToLower(string(norm.NFC.Bytes([]byte(name))))

	for i, v := range a.rows {
		if v.comparableField == comparableField {
			a.rows = append(a.rows[:i], a.rows[i+1:]...)
			return nil
		}
	}

	return fmt.Errorf("Found no row with the same name as the provided row")
}

func (a *ArchiveKey) Print(name string) {

	for i, v := range columns {
		if i < 3 {
			fmt.Printf("%-25s ", v)
		} else {
			fmt.Printf("%10s ", v)
		}
	}

	fmt.Println()

	for _, v := range a.rows {

		if name != "" && strings.ToLower(name) != v.comparableField {
			continue
		}

		fmt.Printf("%-25.25s %-25.25s %-25.25s %10v %10v %10v %10v\n",
			v.Name,
			v.Composer,
			v.Arranger,
			v.HasScore,
			v.HasWoodwind,
			v.HasBrass,
			v.HasPercussion)
	}

	fmt.Printf("\nNumber of rows: %d\n", len(a.rows))
}

func (a *ArchiveKey) Save() error {
	if a.file == nil {
		return fmt.Errorf("cannot save a unloaded file")
	}

	return a.save(a.file)
}

func (a *ArchiveKey) SaveAs(fileName string) error {
	if fileName == "" {
		return fmt.Errorf("missing filename")
	}

	file := excelize.NewFile()
	file.Path = fileName
	return a.save(file)
}

func (a *ArchiveKey) save(file *excelize.File) error {

	for i, v := range columns {
		cell := fmt.Sprintf("%s1", excelize.ToAlphaString(i))
		file.SetCellValue("Sheet1", cell, v)
	}

	for i, r := range a.rows {
		rowNumber := i + 2

		file.SetCellFormula("Sheet1", fmt.Sprintf("A%d", rowNumber), fmt.Sprintf("HYPERLINK(\"%s/%s\", \"%s\")", dropboxURL, r.Name, r.Name))
		file.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowNumber), r.Composer)
		file.SetCellValue("Sheet1", fmt.Sprintf("C%d", rowNumber), r.Arranger)
		file.SetCellValue("Sheet1", fmt.Sprintf("D%d", rowNumber), toString(r.HasScore))
		file.SetCellValue("Sheet1", fmt.Sprintf("E%d", rowNumber), toString(r.HasWoodwind))
		file.SetCellValue("Sheet1", fmt.Sprintf("F%d", rowNumber), toString(r.HasBrass))
		file.SetCellValue("Sheet1", fmt.Sprintf("G%d", rowNumber), toString(r.HasPercussion))
	}


	// Freeze first (header) row
	file.SetPanes("Sheet1",
		fmt.Sprintf(`{"freeze":true,"split":false,"x_split":0,"y_split":1,"top_left_cell":"A2","active_pane":"bottomLeft","panes":[{"sqref":"A2:H%d","active_cell":"A2","pane":"bottomLeft"}]}`, len(a.rows)+1))

	// Format as table
	file.AddTable("Sheet1",
		"A1",
		fmt.Sprintf("G%d", len(a.rows)+1),
		`{"table_style":"TableStyleMedium2", "show_first_column":false,"show_last_column":false,"show_row_stripes":false,"show_column_stripes":false}`)

	// Column widths
	file.SetColWidth("Sheet1", "A", "A", 40)
	file.SetColWidth("Sheet1", "B", "C", 30)
	file.SetColWidth("Sheet1", "D", "H", 15)

	return file.Save()
}

func escape(s string) string {
	return strings.Replace(s, " ", "%20", -1)
}

func (a *ArchiveKey) Load() error {

	file, err := excelize.OpenFile(a.fileName)
	if err != nil {
		return err
	}

	a.file = file

	rows := file.GetRows("Sheet1")
	if len(rows) == 0 {
		return nil
	}

	for i := range rows[1:] {
		rowNum := i + 2

		name, err := nameFromHyperlink(file.GetCellFormula("Sheet1", fmt.Sprintf("A%d", rowNum)))
		if err != nil {
			fmt.Printf("Failed to read name form hyperlink in cell A%d, %v\n", rowNum, err)
			continue
		}

		a.AddRow(name,
			file.GetCellValue("Sheet1", fmt.Sprintf("B%d", rowNum)),
			file.GetCellValue("Sheet1", fmt.Sprintf("C%d", rowNum)),

			toBool(file.GetCellValue("Sheet1", fmt.Sprintf("D%d", rowNum))),
			toBool(file.GetCellValue("Sheet1", fmt.Sprintf("E%d", rowNum))),
			toBool(file.GetCellValue("Sheet1", fmt.Sprintf("F%d", rowNum))),
			toBool(file.GetCellValue("Sheet1", fmt.Sprintf("G%d", rowNum))),
		)
	}

	return nil
}

func nameFromHyperlink(s string) (string, error) {
	if s == "" {
		return "", fmt.Errorf("Empty string")
	}

	hyperlinkFormula := regexp.MustCompile(".*?\"(.*)\",.*?\"(.*)\"")
	matches := hyperlinkFormula.FindStringSubmatch(s)
	if len(matches) != 3 {
		return "", fmt.Errorf("failed to extract formula from: %s", s)
	}

	return matches[2], nil
}

func toBool(s string) bool {
	if s == "Ja" {
		return true
	}

	return false
}

func toString(b bool) string {
	if b {
		return "Ja"
	}
	return "Nei"
}

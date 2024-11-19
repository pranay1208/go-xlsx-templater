package xlst

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/aymerick/raymond"
	xlsx "github.com/tealeg/xlsx/v3"
)

var (
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rangeRgx    = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
)

var ErrEndIterationEarly = errors.New("error to end iteration early")

// Xlst Represents template struct
type Xlst struct {
	file   *xlsx.File
	report *xlsx.File
}

// Options for render has only one property WrapTextInAllCells for wrapping text
type Options struct {
	WrapTextInAllCells bool
}

// New creates new Xlst struct and returns pointer to it
func New() *Xlst {
	return &Xlst{}
}

// NewFromBinary creates new Xlst struct puts binary tempate into and returns pointer to it
func NewFromBinary(content []byte) (*Xlst, error) {
	file, err := xlsx.OpenBinary(content)
	if err != nil {
		return nil, err
	}

	res := &Xlst{file: file}
	return res, nil
}

// Render renders report and stores it in a struct
func (m *Xlst) Render(in interface{}) error {
	// return errors.New("abcde")
	return m.RenderWithOptions(in, nil)
}

// RenderWithOptions renders report with options provided and stores it in a struct
func (m *Xlst) RenderWithOptions(in interface{}, options *Options) error {
	if options == nil {
		options = new(Options)
	}
	report := xlsx.NewFile()
	for si, sheet := range m.file.Sheets {
		ctx := getCtx(in, si)
		report.AddSheet(sheet.Name)
		cloneSheet(sheet, report.Sheets[si])

		err := renderRows(report.Sheets[si], sheet, 0, sheet.MaxRow, ctx, options)
		if err != nil {
			return err
		}

		sheet.Cols.ForEach(func(_ int, col *xlsx.Col) {
			report.Sheets[si].Cols.Add(col)
		})

	}
	m.report = report

	return nil
}

// ReadTemplate reads template from disk and stores it in a struct
func (m *Xlst) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	m.file = file
	return nil
}

// Save saves generated report to disk
func (m *Xlst) Save(path string) error {
	if m.report == nil {
		return errors.New("report was not generated")
	}
	return m.report.Save(path)
}

// Write writes generated report to provided writer
func (m *Xlst) Write(writer io.Writer) error {
	if m.report == nil {
		return errors.New("report was not generated")
	}
	return m.report.Write(writer)
}

func renderRows(sheet *xlsx.Sheet, template *xlsx.Sheet, startIndex int, endIndex int, ctx map[string]interface{}, options *Options) error {
	for ri := startIndex; ri < endIndex; ri++ {
		row, err := template.Row(ri)
		if err != nil {
			return errors.Join(fmt.Errorf("error reading row %d", ri), err)
		}

		rangeProp := getRangeProp(row)
		if rangeProp != "" {
			ri++

			rangeEndIndex, err := getRangeEndIndex(template, ri, endIndex)
			if err != nil {
				return err
			}
			if rangeEndIndex == -1 {
				return fmt.Errorf("end of range %q not found", rangeProp)
			}

			rangeCtx := getRangeCtx(ctx, rangeProp)
			if rangeCtx == nil {
				return fmt.Errorf("not expected context property for range %q", rangeProp)
			}

			for idx := range rangeCtx {
				localCtx := mergeCtx(rangeCtx[idx], ctx)
				err := renderRows(sheet, template, ri, rangeEndIndex, localCtx, options)
				if err != nil {
					return err
				}
			}

			ri = rangeEndIndex

			continue
		}

		prop := getListProp(row)
		if prop == "" {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			err := renderRow(newRow, ctx)
			if err != nil {
				return err
			}
			continue
		}

		if !isArray(ctx, prop) {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			err := renderRow(newRow, ctx)
			if err != nil {
				return err
			}
			continue
		}

		arr := reflect.ValueOf(ctx[prop])
		arrBackup := ctx[prop]
		for i := 0; i < arr.Len(); i++ {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			ctx[prop] = arr.Index(i).Interface()
			err := renderRow(newRow, ctx)
			if err != nil {
				return err
			}
		}
		ctx[prop] = arrBackup
	}

	return nil
}

func cloneCell(from, to *xlsx.Cell, options *Options) {
	to.Value = from.Value
	style := from.GetStyle()
	if options.WrapTextInAllCells {
		style.Alignment.WrapText = true
	}
	to.SetStyle(style)
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt

	if from.Formula() != "" {
		to.SetFormula(from.Formula())
	}
}

func cloneRow(from, to *xlsx.Row, options *Options) {
	if from.GetHeight() != 0 {
		to.SetHeight(from.GetHeight())
	}

	from.ForEachCell(func(c *xlsx.Cell) error {
		newCell := to.AddCell()
		cloneCell(c, newCell, options)
		return nil
	})
}

func renderCell(cell *xlsx.Cell, ctx interface{}) error {
	tpl := strings.Replace(cell.Value, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
	template, err := raymond.Parse(tpl)
	if err != nil {
		return err
	}
	out, err := template.Exec(ctx)
	if err != nil {
		return err
	}

	cellFormat := cell.GetNumberFormat()

	ci, err_ci := strconv.Atoi(out)
	cf, err_cf := strconv.ParseFloat(out, 64)
	ctd, err_ctd := time.Parse(time.RFC3339, out)

	if cell.Formula() != "" {
		// nothing to do, keep original cell state
	} else if err_ci == nil {
		cell.SetInt(ci)
		cell.NumFmt = cellFormat
	} else if err_cf == nil {
		cell.SetFloat(cf)
		cell.NumFmt = cellFormat
	} else if err_ctd == nil {
		cell.SetDateTime(ctd)
	} else {
		cell.SetValue(out)
	}

	return nil
}

func cloneSheet(from, to *xlsx.Sheet) {
	from.Cols.ForEach(func(_ int, col *xlsx.Col) {
		newCol := xlsx.Col{}
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max
		to.Cols.Add(&newCol)
	})
}

func getCtx(in interface{}, i int) map[string]interface{} {
	if ctx, ok := in.(map[string]interface{}); ok {
		return ctx
	}
	if ctxSlice, ok := in.([]interface{}); ok {
		if len(ctxSlice) > i {
			_ctx := ctxSlice[i]
			if ctx, ok := _ctx.(map[string]interface{}); ok {
				return ctx
			}
		}
		return nil
	}
	return nil
}

func getRangeCtx(ctx map[string]interface{}, prop string) []map[string]interface{} {
	val, ok := ctx[prop]
	if !ok {
		return nil
	}

	if propCtx, ok := val.([]map[string]interface{}); ok {
		return propCtx
	}

	return nil
}

func mergeCtx(local, global map[string]interface{}) map[string]interface{} {
	ctx := make(map[string]interface{})

	for k, v := range global {
		ctx[k] = v
	}

	for k, v := range local {
		ctx[k] = v
	}

	return ctx
}

func isArray(in map[string]interface{}, prop string) bool {
	val, ok := in[prop]
	if !ok {
		return false
	}
	switch reflect.TypeOf(val).Kind() {
	case reflect.Array, reflect.Slice:
		return true
	}
	return false
}

func getListProp(in *xlsx.Row) string {
	matchFoundInFn := ""
	in.ForEachCell(func(cell *xlsx.Cell) error {
		if match := rgx.FindAllStringSubmatch(cell.Value, -1); match != nil {
			matchFoundInFn = match[0][1]
			return ErrEndIterationEarly
		}
		return nil
	}, xlsx.SkipEmptyCells)

	return matchFoundInFn
}

func getRangeProp(in *xlsx.Row) string {
	cell := in.GetCell(0)
	match := rangeRgx.FindAllStringSubmatch(cell.Value, -1)
	if match != nil {
		return match[0][1]
	}
	return ""
}

func getRangeEndIndex(sheet *xlsx.Sheet, fromIndex, toIndex int) (int, error) {
	var nesting int
	for idx := fromIndex; idx < toIndex; idx++ {
		row, err := sheet.Row(idx)
		if err != nil {
			return -1, errors.Join(fmt.Errorf("error reading row %d", idx), err)
		}

		cell := row.GetCell(0)
		if rangeEndRgx.MatchString(cell.Value) {
			if nesting == 0 {
				return idx, nil
			}

			nesting--
			continue
		}

		if rangeRgx.MatchString(cell.Value) {
			nesting++
		}
	}

	return -1, nil
}

func renderRow(in *xlsx.Row, ctx interface{}) error {
	return in.ForEachCell(func(cell *xlsx.Cell) error {
		return renderCell(cell, ctx)
	})
}

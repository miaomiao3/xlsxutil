package xlsxutil

import (
	"errors"
	"github.com/miaomiao3/xlsx" // i fork this repo to enable SetType
	"reflect"
	"strings"
)

var (
	errInputType = errors.New("wants pointer of slice, slice contains struct element")
)

// dump data to file
// data: ptr of slice, slice element should be a struct
// sep: separator of csv line. be careful to avoid value conflict
func XlsDump(file *xlsx.File, sheetName string, data interface{}) error {
	dataValue := reflect.ValueOf(data)
	if dataValue.Kind() != reflect.Slice {
		return errors.New("do not support non slice data")
	}
	sliceLen := dataValue.Len()
	if sliceLen == 0 {
		return errors.New("empty slice")
	}

	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return err

	}
	var optionMap map[string]*xlsOption
	for i := 0; i < sliceLen; i++ {

		itemValue := dataValue.Index(i)
		if itemValue.Kind() == reflect.Ptr {
			itemValue = itemValue.Elem()
		}
		if i == 0 {
			row := sheet.AddRow()
			optionMap, err = addHeaderRow(itemValue, func(str string, kind reflect.Kind) {
				cell := row.AddCell()
				cell.Value = str
				cell.SetType(xlsx.CellTypeString)
			})
			if err != nil {
				return err
			}
			continue
		}
		row := sheet.AddRow()
		err = addRow(itemValue, optionMap, func(str string, kind reflect.Kind) {
			cell := row.AddCell()
			cell.Value = str
			switch kind {
			case reflect.Interface, reflect.Map, reflect.Array, reflect.Slice, reflect.Complex64,
				reflect.Complex128, reflect.UnsafePointer, reflect.Chan, reflect.Func,
				reflect.String:
				cell.SetType(xlsx.CellTypeString)

			case reflect.Int8, reflect.Uint8, reflect.Int16, reflect.Uint16, reflect.Int32, reflect.Uint32,
				reflect.Int64, reflect.Uint64, reflect.Int, reflect.Uint, reflect.Float64, reflect.Float32:
				cell.SetType(xlsx.CellTypeNumeric)

			}
		})
		if err != nil {
			return err
		}
	}
	return err
}

// load one sheet to slice
func XlsLoad(file *xlsx.File, sheetName string, data interface{}) error {
	sheet, ok := file.Sheet[sheetName]
	if !ok {
		return errors.New("sheetName not found")
	}

	dataType, sliceValue, isElementPtr, err := validateDataInput(data)
	if err != nil {
		return err
	}

	dataValue := reflect.New(*dataType).Elem()
	_, optionMap := getStructOptions(dataValue)

	// column index ->  column cell string
	headerMap := make(map[int]string)

	for rowIndex, row := range sheet.Rows {
		if rowIndex == 0 {
			for columnIndex, cell := range row.Cells {
				if len(cell.String()) == 0 { // if value is empty, ignore
					continue
				}
				headerMap[columnIndex] = cell.String()
			}
			continue
		}

		// check if this row is empty
		isRowEmpty := true
		for _, cell := range row.Cells {
			if len(strings.TrimSpace(cell.String())) > 0 {
				isRowEmpty = false
				break
			}
		}

		if isRowEmpty {
			break
		}

		valueMap := make(map[string]string)
		for columnIndex, cell := range row.Cells {
			if len(headerMap[columnIndex]) == 0 { // if head is empty, ignore
				continue
			}
			valueMap[headerMap[columnIndex]] = cell.String()
		}

		addElement(*sliceValue, *dataType, isElementPtr, valueMap, optionMap)
	}

	return nil
}

package xlsxutil

import (
	"fmt"
	"testing"

	"github.com/miaomiao3/xlsx"
	. "github.com/smartystreets/goconvey/convey"
)

const (
	xlsxFilePath  = "example/xlsx/people.xlsx"
	xlsxFileSheet = "persons"
)

func TestXlsDump(t *testing.T) {
	Convey("TestXlsDump", t, func() {
		persons := prepareTestData()
		file := xlsx.NewFile()
		err := XlsDump(file, xlsxFileSheet, persons)
		So(err, ShouldEqual, nil)
		err = file.Save(xlsxFilePath)
		So(err, ShouldEqual, nil)
	})
}

func TestXlsxBind(t *testing.T) {
	Convey("TestXlsxBind", t, func() {
		persons := make([]*Person, 0)
		file, err := xlsx.OpenFile(xlsxFilePath)
		So(err, ShouldEqual, nil)

		err = XlsLoad(file, xlsxFileSheet, &persons)
		fmt.Println("persons:", persons)
		So(err, ShouldEqual, nil)
		So(len(persons), ShouldEqual, 4)
		So(persons[3].Name, ShouldEqual, "n-4")
	})
}

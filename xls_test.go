package xlsxutil

import (
	"github.com/miaomiao3/xlsx"
	. "github.com/smartystreets/goconvey/convey"
	"testing"
)


const (
	xlsxFilePath = "example/xlsx/prople.xlsx"
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

func TestXlsxBindByYamlTag(t *testing.T) {
	Convey("TestXlsxBindByYamlTag", t, func() {
		persons := make([]*Person, 0)
		file, err := xlsx.OpenFile(xlsxFilePath)
		So(err, ShouldEqual, nil)

		err = XlsBindByYamlTag(file, xlsxFileSheet, &persons)
		So(err, ShouldEqual, nil)
		So(len(persons), ShouldEqual, 5)
		So(persons[4].Name, ShouldEqual, "n-4")
	})
}

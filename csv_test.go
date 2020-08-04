package xlsxutil

import (
	"fmt"
	. "github.com/smartystreets/goconvey/convey"
	"testing"
)

type Person struct {
	Name  string  `xls:"name",yaml:"name"`
	Money float64 `xls:"money,precision:10",yaml:"money"`
	Age   int     `xls:"age",yaml:"age"`
	Edu   `xls:",inline"`
}

type Edu struct {
	School  string `xls:"school",yaml:"school"`
	Address string `xls:"address",yaml:"address"`
}

const (
	personCsvStr = `name,money,age,school,address
n-0,1.2345678900,20,school-0,hali-0
n-1,1.2345678900,21,school-1,hali-1
n-2,1.2345678900,22,school-2,hali-2
n-3,1.2345678900,23,school-3,hali-3
n-4,1.2345678900,24,school-4,hali-4
`
)

func prepareTestData()([]*Person){
	persons := make([]*Person, 0, 5)
	for i := 0; i < 5; i++ {
		newPerson := &Person{
			Name:  fmt.Sprintf("n-%d", i),
			Age:   i + 20,
			Money: 1.23456789,
		}
		newPerson.School = fmt.Sprintf("school-%d", i)
		newPerson.Address = fmt.Sprintf("hali-%d", i)
		persons = append(persons, newPerson)
	}
	return persons
}


func TestCsvDump(t *testing.T) {
	Convey("CsvDump", t, func() {
		persons := prepareTestData()
		buf, err := CsvDump(",", persons)

		So(err, ShouldEqual, nil)
		So(buf.String(), ShouldEqual, personCsvStr)
	})
}

func TestCsvBindByYamlTag(t *testing.T) {
	Convey("TestCsvBindByYamlTag", t, func() {
		persons := make([]*Person, 0)
		err := CsvBindByYamlTag("./example/csv/people.csv", ",", &persons)

		So(err, ShouldEqual, nil)
		So(len(persons), ShouldEqual, 5)
		So(persons[4].Name, ShouldEqual, "n-4")
	})
}

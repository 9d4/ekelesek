package main

import (
	"fmt"

	"github.com/9d4/ekelesek/eksel"
	"github.com/xuri/excelize/v2"
)

// the key is the column name in xlsx file
// the value is the lookup tag in the target struct
var lookupMap = map[string]string{
	"Name":      "name",
	"Age":       "age",
	"Birth Day": "birthday",
}

type Student struct {
	Name     string `lookup:"name"`
	Age      int    `lookup:"age"`
	BirthDay string `lookup:"birthday"`
}

func main() {
	f, err := excelize.OpenFile("./data.xlsx")
	checkErr(err)
	rows, err := f.Rows("Sheet1")
	checkErr(err)

	var students []Student
	checkErr(eksel.Parse(rows, lookupMap, &students))

	for _, student := range students {
		fmt.Printf("Name: %s\nAge: %d\nBirthday: %s\n\n", student.Name, student.Age, student.BirthDay)
	}
}

func checkErr(err error) {
	if err != nil {
		panic(err)
	}
}

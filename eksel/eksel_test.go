package eksel

import (
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"log"
	"testing"
	"time"
)

func TestParse(t *testing.T) {
	file, err := excelize.OpenFile("./testdata/data-50.xlsx")
	assert.NoError(t, err, "unable to open file")
	rows, err := file.Rows("Sheet1")
	assert.NoError(t, err)

	headerMap := map[string]string{
		"Name":             "name",
		"Phone":            "phone",
		"Email":            "email",
		"Age":              "age",
		"Address":          "address",
		"Favourite Number": "fav_num",
		"Country":          "country",
	}
	type Data struct {
		Name    string `lookup:"name"`
		Phone   string `lookup:"phone"`
		Email   string `lookup:"email"`
		Age     int    `lookup:"age"`
		Address string `lookup:"address"`
		FavNum  int    `lookup:"fav_num"`
		Country string `lookup:"country"`
	}
	d := []Data{}

	err = Parse(rows, headerMap, &d)
	assert.NoError(t, err)

	out := []Data{{Name: "Patience Whitfield", Phone: "(954) 845-3772", Email: "montes.nascetur@hotmail.org", Age: 19, Address: "Ap #726-6321 Aliquam Street", FavNum: 1, Country: "Vietnam"}, {Name: "Brennan Collins", Phone: "(232) 454-4524", Email: "tempor@icloud.edu", Age: 23, Address: "867-988 Sed St.", FavNum: 15, Country: "Mexico"}}
	assert.Equal(t, out, d)
}

func TestParse2(t *testing.T) {
	file, err := excelize.OpenFile("./testdata/data-unsorted-header.xlsx")
	assert.NoError(t, err, "unable to open file")
	rows, err := file.Rows("Sheet1")
	assert.NoError(t, err)

	headerMap := map[string]string{
		"Name":     "name",
		"Age":      "age",
		"Birthday": "birthday",
	}
	type Data struct {
		Name     string `lookup:"name"`
		Age      int    `lookup:"age"`
		Birthday string `lookup:"birthday"`
	}
	d := []Data{}

	err = Parse(rows, headerMap, &d)
	assert.NoError(t, err)

	expect := []Data{{Name: "Kaye Goff", Age: 26, Birthday: "April"}, {Name: "Adrienne Kirby", Age: 22, Birthday: "May"}, {Name: "John", Age: 27, Birthday: "May"}}
	assert.Equal(t, expect, d)
}

func TestParse_time(t *testing.T) {
	rows := openT(t, "./testdata/data-time.xlsx")
	headerMap := map[string]string{
		"Time Start": "ts",
		"Time End":   "te",
		"Date":       "date",
	}
	type Data struct {
		Start time.Time `lookup:"ts"`
		End   time.Time `lookup:"te"`
		Date  time.Time `lookup:"date"`
	}
	var d []Data
	assert.NoError(t, Parse(rows, headerMap, &d))
	log.Printf("%#v", d)
}

func BenchmarkParse_1000(b *testing.B) {
	rows := open(b, "./testdata/data-1000.xlsx")
	headerMap := map[string]string{
		"Name":             "name",
		"Phone":            "phone",
		"Email":            "email",
		"Age":              "age",
		"Address":          "address",
		"Favourite Number": "fav_num",
		"Country":          "country",
	}
	type Data struct {
		Name    string `lookup:"name"`
		Phone   string `lookup:"phone"`
		Email   string `lookup:"email"`
		Age     int    `lookup:"age"`
		Address string `lookup:"address"`
		FavNum  int    `lookup:"fav_num"`
		Country string `lookup:"country"`
	}

	for i := 0; i < b.N; i++ {
		var data []Data
		assert.NoError(b, Parse(rows, headerMap, &data))
	}
}

func open(b *testing.B, filename string) *excelize.Rows {
	b.Helper()
	file, err := excelize.OpenFile(filename)
	assert.NoError(b, err, "unable to open file")
	rows, err := file.Rows("Sheet1")
	assert.NoError(b, err)
	return rows
}

func openT(t *testing.T, filename string) *excelize.Rows {
	t.Helper()
	file, err := excelize.OpenFile(filename)
	assert.NoError(t, err, "unable to open file")
	rows, err := file.Rows("Sheet1")
	assert.NoError(t, err)
	return rows
}

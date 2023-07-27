package eksel

import (
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"testing"
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

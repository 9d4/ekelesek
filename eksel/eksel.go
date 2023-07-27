package eksel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

// Parse transforms rows into desired list of struct.
// header map[string]string is mapping the header row to dest.
//
// Example:
//
//	header := map[string]string{
//	// the key is the Text in cell of header
//	// the value is the field tag to dest
//	 	"Name": 			"name"
//		"Favourite Animal": "fav_animal"
//	}
//
//
//	type Person struct {
//		Name 		string `lookup:"name"`
//		FavAnimal 	string `lookup:"fav_animal"`
//	}
//
//	dest := []Person{}
func Parse(rows *excelize.Rows, header map[string]string, dest interface{}) error {
	// Ensure that dest is a slice of structs
	destValue := reflect.ValueOf(dest)
	if destValue.Kind() != reflect.Ptr || destValue.Elem().Kind() != reflect.Slice {
		return errors.New("dest must be a pointer to a slice")
	}

	// Get the type of the slice elements
	elemType := destValue.Elem().Type().Elem()
	if elemType.Kind() != reflect.Struct {
		return errors.New("dest must be a pointer to a slice of structs")
	}

	fieldMapIndex := map[int]string{}
	// Map header keys to struct field names using tag "lookup"
	fieldMap := make(map[string]string)
	for i := 0; i < elemType.NumField(); i++ {
		field := elemType.Field(i)
		lookupTag := field.Tag.Get("lookup")
		if lookupTag != "" {
			fieldMap[lookupTag] = field.Name
		}
	}

	// header[idx]
	// loop item, then item.

	// Read Excel rows and populate the slice of structs
	i := 0
	for rows.Next() {
		// Create a new instance of the struct
		item := reflect.New(elemType).Elem()

		// Get the row values from Excel
		cells, err := rows.Columns()
		if err != nil {
			return err
		}

		// Map header index
		if i == 0 {
			for idx, cell := range cells {
				lookupKey, exists := header[cell]
				if exists {
					fieldMapIndex[idx] = fieldMap[lookupKey]
				}
			}
			i++
			continue
		}

		for idx, fieldName := range fieldMapIndex {
			field := item.FieldByName(fieldName)
			if !field.IsValid() {
				return errors.New("invalid field: " + field.String())
			}
			if field.Kind() == reflect.String {
				field.SetString(cells[idx])
			}
			if field.Kind() == reflect.Int {
				integer, _ := strconv.Atoi(cells[idx])
				field.SetInt(int64(integer))
			}
		}

		destValue.Elem().Set(reflect.Append(destValue.Elem(), item))
	}

	return nil
}

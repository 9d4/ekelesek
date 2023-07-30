package eksel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
	"strings"
	"time"
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

	// Case-insensitive. copy original to prevent reference bug
	copyHeader := make(map[string]string)
	for s, s2 := range header {
		copyHeader[strings.ToLower(s)] = s2
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
				lookupKey, exists := copyHeader[strings.ToLower(cell)]
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

			cellValue := ""
			if idx < len(cells) {
				cellValue = cells[idx]
			}

			switch field.Kind() {
			case reflect.String:
				field.SetString(cellValue)
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				intVal, err := strconv.ParseInt(cellValue, 10, 64)
				if err == nil {
					field.SetInt(intVal)
				}
			case reflect.Float32, reflect.Float64:
				floatVal, err := strconv.ParseFloat(cellValue, 64)
				if err == nil {
					field.SetFloat(floatVal)
				}
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				uintVal, err := strconv.ParseUint(cellValue, 10, 64)
				if err == nil {
					field.SetUint(uintVal)
				}
			case reflect.Bool:
				boolVal, err := strconv.ParseBool(strings.ToLower(cellValue))
				if err == nil {
					field.SetBool(boolVal)
				}
			case reflect.Struct:
				f := reflect.Indirect(item).FieldByName(fieldName)
				fi := f.Interface()

				switch fi.(type) {
				case time.Time:
					var timeVal time.Time
					var err error

					dateLayouts := []string{
						"2006-01-02",
						"02/01/06",
						"02/01/2006",
					}
					timeLayouts := []string{
						"15:04:05",
						"15:04",
					}
					combinedLayouts := []string{
						"02/01/2006 15:04",
						"02/01/2006 15:04:05",
						"2006-01-02 15:04",
						"2006-01-02 15:04:05",
						"02/01/06 15:04",
						"02/01/06 15:04:05",
					}

					for _, layout := range combinedLayouts {
						timeVal, err = time.Parse(layout, cellValue)
						if err == nil {
							break
						}
					}

					if err != nil {
						for _, layout := range dateLayouts {
							timeVal, err = time.Parse(layout, cellValue)
							if err == nil {
								break
							}
						}
					}

					if err != nil {
						for _, layout := range timeLayouts {
							timeVal, err = time.Parse(layout, cellValue)
							if err == nil {
								break
							}
						}
					}

					field.Set(reflect.ValueOf(timeVal))
				}
			default:
				return errors.New("unsupported field type: " + field.Type().Name())
			}
		}

		destValue.Elem().Set(reflect.Append(destValue.Elem(), item))
	}

	return nil
}

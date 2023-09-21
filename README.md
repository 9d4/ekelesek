# ekelesek

Extract xlsx to struct with [excelize](github.com/xuri/excelize).

## Usage

| Age | Name           | Birth Day |
| --- | -------------- | --------- |
| 26  | Kaye Goff      | April     |
| 22  | Adrienne Kirby | May       |
| 27  | John           | May       |

Make new type struct.

```go
type Student struct {
	Name     string `lookup:"name"`
	Age      int    `lookup:"age"`
	BirthDay string `lookup:"birthday"`
}
```

Map the xlsx column.

```go
var lookupMap = map[string]string{
	"Name":      "name",
	"Age":       "age",
	"Birth Day": "birthday",
}
```

The key is the column name in the table and the value is the lookup key for Student struct.

Full example see in example directory or below code.

```go
package main

import (
	"fmt"

	"github.com/9d4/ekelesek/eksel"
	"github.com/xuri/excelize/v2"
)

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
	f, _ := excelize.OpenFile("./data.xlsx")
	rows, _ := f.Rows("Sheet1")

	var students []Student
	eksel.Parse(rows, lookupMap, &students)

	for _, student := range students {
		fmt.Printf("Name: %s\nAge: %d\nBirthday: %s\n\n", student.Name, student.Age, student.BirthDay)
	}
}
```

### Supported type
- Int
- Float
- Uint
- Bool
- Time

> Note: The code is not well tested, I only use this for string, int, and datetime type.

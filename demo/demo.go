package main

import (
	"fmt"

	xlst "github.com/Sabaverus/go-xlsx-templater"
)

func main() {
	doc := xlst.New()
	if err := doc.ReadTemplate("./template.xlsx"); err != nil {
		fmt.Printf("Got error reading template: %v", err)
		return
	}
	ctx := map[string]interface{}{
		"name":           "Github User",
		"nameHeader":     "Item name",
		"quantityHeader": "Quantity",
		"items": []map[string]interface{}{
			{
				"name":     "Pen",
				"quantity": 2,
			},
			{
				"name":     "Pencil",
				"quantity": 1,
			},
			{
				"name":     "Condom",
				"quantity": 12,
			},
			{
				"name":     "Beer",
				"quantity": 24,
			},
		},
	}
	err := doc.Render(ctx)
	if err != nil {
		panic(err)
	}
	err = doc.Save("./report.xlsx")
	if err != nil {
		panic(err)
	}
}

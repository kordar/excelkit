package excelutil

import (
	"testing"
)

type TestUser struct {
	ID       int    `json:"id" excel:"ID"`
	Name     string `json:"name" excel:"姓名"`
	Age      int    `json:"age" excel:"年龄"`
	Password string `json:"password" excel:"-"`
	NoExcel  string `json:"no_excel"`
}

func TestParseColumns(t *testing.T) {
	cols := ParseColumns[TestUser]()
	if len(cols) != 5 {
		t.Fatalf("expected 5 columns, got %d", len(cols))
	}

	if cols[3].Ignore != true {
		t.Errorf("expected Password to be ignored")
	}

	if cols[4].Header != "NoExcel" {
		t.Errorf("expected NoExcel to fallback to field name, got %s", cols[4].Header)
	}
}

func TestGetNestedValue(t *testing.T) {
	m := map[string]any{
		"user": map[string]any{
			"profile": map[string]any{
				"name": "John",
			},
		},
		"simple": "value",
	}

	if v := GetNestedValue(m, "simple"); v != "value" {
		t.Errorf("expected 'value', got %v", v)
	}

	if v := GetNestedValue(m, "user.profile.name"); v != "John" {
		t.Errorf("expected 'John', got %v", v)
	}

	if v := GetNestedValue(m, "user.unknown"); v != nil {
		t.Errorf("expected nil, got %v", v)
	}
}

func TestColumnBuilderAndOrder(t *testing.T) {
	defaultCols := []Column{
		{Key: "a", Header: "A", Order: 0},
		{Key: "b", Header: "B", Order: 0},
	}

	customCols := []Column{
		Col("c").Header("C").Order(-1).Build(),
		Col("d").Header("D").Order(1).Build(),
	}

	merged := MergeColumns(defaultCols, customCols)

	if len(merged) != 4 {
		t.Fatalf("expected 4 columns, got %d", len(merged))
	}

	// Expected order: c (order -1), a (order 0), b (order 0), d (order 1)
	expectedKeys := []string{"c", "a", "b", "d"}
	for i, col := range merged {
		if col.Key != expectedKeys[i] {
			t.Errorf("expected column at index %d to be %s, got %s", i, expectedKeys[i], col.Key)
		}
	}
}

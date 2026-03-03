package excelkit

import "github.com/xuri/excelize/v2"

// --------------------
// 样式封装
// --------------------
type Style struct {
	Font      *excelize.Font
	Fill      *excelize.Fill
	Border    []excelize.Border
	Alignment *excelize.Alignment
}

func (s *Style) Build(f *excelize.File) (int, error) {
	// 构建 excelize 的 Style
	st := &excelize.Style{
		Font:      s.Font,
		Border:    s.Border,
		Alignment: s.Alignment,
	}
	if s.Fill != nil {
		st.Fill = *s.Fill
	}
	return f.NewStyle(st)
}

func RedStyle() *Style {
	return &Style{Font: &excelize.Font{Color: "FF0000"}}
}

func GreenStyle() *Style {
	return &Style{Font: &excelize.Font{Color: "00AA00"}}
}

func BlueSummaryStyle() *Style {
	return &Style{
		Font: &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill: &excelize.Fill{Type: "pattern", Color: []string{"0000FF"}, Pattern: 1},
	}
}

func BorderedStyle() *Style {
	return &Style{
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	}
}

func BorderedHeaderStyle() *Style {
	return &Style{
		Font: &excelize.Font{Bold: true},
		Fill: &excelize.Fill{Type: "pattern", Color: []string{"DDDDDD"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
	}
}

func BlueSummaryBorderedStyle() *Style {
	return &Style{
		Font: &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill: &excelize.Fill{Type: "pattern", Color: []string{"0000FF"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	}
}

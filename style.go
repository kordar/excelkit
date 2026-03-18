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

func TableHeaderBlueStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"1F4E79"}, Pattern: 1},
		Border:    borderThin("1F4E79"),
		Alignment: alignCenterWrap(),
	}
}

func TableBodyBlueStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Color: "1F1F1F"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
		Border:    borderThin("D9D9D9"),
		Alignment: alignLeftWrap(),
	}
}

func TableHeaderDarkStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"333333"}, Pattern: 1},
		Border:    borderThin("333333"),
		Alignment: alignCenterWrap(),
	}
}

func TableBodyDarkStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Color: "1F1F1F"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
		Border:    borderThin("BFBFBF"),
		Alignment: alignLeftWrap(),
	}
}

func TableHeaderLightStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Bold: true, Color: "1F1F1F"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"F2F2F2"}, Pattern: 1},
		Border:    borderThin("BFBFBF"),
		Alignment: alignCenterWrap(),
	}
}

func TableBodyLightStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Color: "1F1F1F"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
		Border:    borderThin("E0E0E0"),
		Alignment: alignLeftWrap(),
	}
}

func TableHeaderGreenStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"2E7D32"}, Pattern: 1},
		Border:    borderThin("2E7D32"),
		Alignment: alignCenterWrap(),
	}
}

func TableBodyGreenStyle() *Style {
	return &Style{
		Font:      &excelize.Font{Color: "1F1F1F"},
		Fill:      &excelize.Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
		Border:    borderThin("C8E6C9"),
		Alignment: alignLeftWrap(),
	}
}

func borderThin(color string) []excelize.Border {
	return []excelize.Border{
		{Type: "top", Color: color, Style: 1},
		{Type: "bottom", Color: color, Style: 1},
		{Type: "left", Color: color, Style: 1},
		{Type: "right", Color: color, Style: 1},
	}
}

func alignCenterWrap() *excelize.Alignment {
	return &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true}
}

func alignLeftWrap() *excelize.Alignment {
	return &excelize.Alignment{Horizontal: "left", Vertical: "center", WrapText: true}
}

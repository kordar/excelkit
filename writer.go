package excelkit

import "github.com/xuri/excelize/v2"

// --------------------
// ExcelWriter
// --------------------
type ExcelWriter[T any] struct {
	File           *excelize.File
	Schema         SheetSchema[T]
	Source         DataSource[T]
	UseStream      bool
	Interceptors   []RowInterceptor[T]
	sw             *excelize.StreamWriter
	headerStyleID  int
	defaultStyleID int
}

func (w *ExcelWriter[T]) Write() error {
	f := w.File
	sheet := w.Schema.Name
	var err error
	if w.UseStream {
		if idx, e := f.GetSheetIndex(sheet); e != nil || idx == -1 {
			_, _ = f.NewSheet(sheet)
		}
		w.sw, err = f.NewStreamWriter(sheet)
		if err != nil {
			return err
		}
	} else {
		if idx, e := f.GetSheetIndex(sheet); e != nil || idx == -1 {
			f.NewSheet(sheet)
		}
	}

	if w.Schema.HeaderStyle != nil {
		sid, err := w.Schema.HeaderStyle.Build(f)
		if err != nil {
			return err
		}
		w.headerStyleID = sid
	}
	if w.Schema.DefaultStyle != nil {
		sid, err := w.Schema.DefaultStyle.Build(f)
		if err != nil {
			return err
		}
		w.defaultStyleID = sid
	}

	// 写 header
	header := make([]interface{}, len(w.Schema.Columns))
	for i, col := range w.Schema.Columns {
		if w.headerStyleID > 0 {
			header[i] = excelize.Cell{Value: col.Header, StyleID: w.headerStyleID}
		} else {
			header[i] = col.Header
		}
	}
	if w.UseStream {
		_ = w.sw.SetRow("A1", header)
	} else {
		for i, val := range header {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue(sheet, cell, val)
			if w.headerStyleID > 0 {
				f.SetCellStyle(sheet, cell, cell, w.headerStyleID)
			}
		}
	}

	rowIdx := 2
	for {
		row, ok, err := w.Source.Next()
		if err != nil || !ok {
			break
		}

		write := true
		extraRows := make([]T, 0)
		mergeRanges := make([]MergeRange[T], 0)

		for _, interceptor := range w.Interceptors {
			wFlag, extra, merge := interceptor(row, rowIdx)
			if !wFlag {
				write = false
			}
			if extra != nil {
				extraRows = append(extraRows, extra...)
			}
			if merge != nil {
				mergeRanges = append(mergeRanges, merge...)
			}
		}

		rowStyleOverrides := map[int]map[int]int{}
		if w.UseStream {
			for _, m := range mergeRanges {
				if m.Style == nil {
					continue
				}
				styleID, err := m.Style.Build(w.File)
				if err != nil {
					return err
				}
				for r := m.StartRow; r <= m.EndRow; r++ {
					rm := rowStyleOverrides[r]
					if rm == nil {
						rm = map[int]int{}
						rowStyleOverrides[r] = rm
					}
					for c := m.StartCol; c <= m.EndCol; c++ {
						rm[c] = styleID
					}
				}
			}
		}

		// 执行合并
		for _, m := range mergeRanges {
			startCell, _ := excelize.CoordinatesToCellName(m.StartCol, m.StartRow)
			endCell, _ := excelize.CoordinatesToCellName(m.EndCol, m.EndRow)
			if w.UseStream && w.sw != nil {
				_ = w.sw.MergeCell(startCell, endCell)
			} else {
				_ = w.File.MergeCell(sheet, startCell, endCell)
			}

			if !w.UseStream && m.Style != nil {
				styleID, _ := m.Style.Build(w.File)
				for r := m.StartRow; r <= m.EndRow; r++ {
					for c := m.StartCol; c <= m.EndCol; c++ {
						cell, _ := excelize.CoordinatesToCellName(c, r)
						w.File.SetCellStyle(sheet, cell, cell, styleID)
					}
				}
			}
		}

		// 写额外行
		for _, er := range extraRows {
			w.writeRow(er, rowIdx, rowStyleOverrides[rowIdx])
			rowIdx++
		}

		if write {
			w.writeRow(row, rowIdx, rowStyleOverrides[rowIdx])
			rowIdx++
		}
	}

	if w.UseStream {
		_ = w.sw.Flush()
	}
	return nil
}

func (w *ExcelWriter[T]) writeRow(row T, rowIdx int, styleOverrides map[int]int) {
	f := w.File
	sheet := w.Schema.Name
	values := make([]interface{}, len(w.Schema.Columns))
	for i, col := range w.Schema.Columns {
		ctx := &CellContext[T]{File: f, Sheet: sheet, RowIndex: rowIdx, ColIndex: i + 1, RowData: row}
		if col.Render != nil {
			_ = col.Render(ctx)
			values[i] = nil
			continue
		}
		values[i] = col.Value(row)

		styleID := 0
		hasStyle := false
		if w.defaultStyleID > 0 {
			styleID = w.defaultStyleID
			hasStyle = true
		}
		if col.Style != nil {
			style := col.Style(row)
			if style != nil {
				sid, _ := style.Build(f)
				styleID = sid
				hasStyle = true
			}
		}
		if styleOverrides != nil {
			if sid, ok := styleOverrides[i+1]; ok {
				styleID = sid
				hasStyle = true
			}
		}

		if hasStyle {
			if w.UseStream {
				values[i] = excelize.Cell{Value: values[i], StyleID: styleID}
			} else {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowIdx)
				f.SetCellStyle(sheet, cell, cell, styleID)
			}
		}
	}
	if w.UseStream {
		cell, _ := excelize.CoordinatesToCellName(1, rowIdx)
		_ = w.sw.SetRow(cell, values)
	} else {
		for i, val := range values {
			cell, _ := excelize.CoordinatesToCellName(i+1, rowIdx)
			f.SetCellValue(sheet, cell, val)
		}
	}
}

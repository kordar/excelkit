package excelutil

import (
	"net/http"
	"net/url"
	"reflect"
	"sort"
	"strings"
	"sync"

	"github.com/kordar/excelkit"
)

// Column 定义一个中间结构（适配 excelkit）
type Column struct {
	Header string
	Key    string
	Ignore bool
	Value  func(map[string]any) any // ⭐ 新增
	Order  int                      // 排序优先级，可选
}

var parseColumnsCache sync.Map

// ParseColumns 从 struct 解析出 Column 定义
func ParseColumns[T any]() []Column {
	var t T
	typ := reflect.TypeOf(t)

	if typ.Kind() == reflect.Ptr {
		typ = typ.Elem()
	}

	if typ.Kind() != reflect.Struct {
		return nil
	}

	// 缓存读取
	cacheKey := typ.String()
	if cached, ok := parseColumnsCache.Load(cacheKey); ok {
		// 返回一个副本，防止外部修改缓存的切片
		cols := cached.([]Column)
		res := make([]Column, len(cols))
		copy(res, cols)
		return res
	}

	var cols []Column

	for i := 0; i < typ.NumField(); i++ {
		f := typ.Field(i)

		if !f.IsExported() {
			continue
		}

		jsonTag := f.Tag.Get("json")
		excelTag := f.Tag.Get("excel")

		if jsonTag == "" || jsonTag == "-" {
			continue
		}

		ignore := false
		if excelTag == "-" {
			ignore = true
		}

		jsonKey := jsonTag
		if idx := strings.IndexByte(jsonTag, ','); idx != -1 {
			jsonKey = jsonTag[:idx]
		}
		if jsonKey == "" {
			continue
		}

		header := excelTag
		if header == "" || header == "-" {
			header = f.Name
		}

		key := jsonKey
		k := key // 修复闭包变量捕获问题
		cols = append(cols, Column{
			Header: header,
			Key:    k,
			Ignore: ignore,
			Value: func(m map[string]any) any {
				return GetNestedValue(m, k)
			},
		})
	}

	parseColumnsCache.Store(cacheKey, cols)

	// 返回一个副本
	res := make([]Column, len(cols))
	copy(res, cols)
	return res
}

// ColumnBuilder 提供流式 API 来构建 Column
type ColumnBuilder struct {
	col Column
}

// Col 初始化一个 ColumnBuilder
func Col(key string) *ColumnBuilder {
	return &ColumnBuilder{
		col: Column{Key: key},
	}
}

// Header 设置表头
func (b *ColumnBuilder) Header(header string) *ColumnBuilder {
	b.col.Header = header
	return b
}

// Format 设置格式化函数
func (b *ColumnBuilder) Format(fn func(map[string]any) any) *ColumnBuilder {
	b.col.Value = fn
	return b
}

// Ignore 设置是否忽略该列
func (b *ColumnBuilder) Ignore(ignore bool) *ColumnBuilder {
	b.col.Ignore = ignore
	return b
}

// Order 设置列的排序优先级
func (b *ColumnBuilder) Order(order int) *ColumnBuilder {
	b.col.Order = order
	return b
}

// Build 最终返回构建好的 Column
func (b *ColumnBuilder) Build() Column {
	return b.col
}

// GetNestedValue supports dot-separated nested fields, e.g. "user.name"
func GetNestedValue(m map[string]any, path string) any {
	if path == "" {
		return nil
	}

	// Fast path for non-nested keys
	if !strings.Contains(path, ".") {
		return m[path]
	}

	parts := strings.Split(path, ".")
	var current any = m

	for _, part := range parts {
		if current == nil {
			return nil
		}

		switch v := current.(type) {
		case map[string]any:
			current = v[part]
		default:
			return nil
		}
	}

	return current
}

func MergeColumns(defaultCols, customCols []Column) []Column {
	// key -> Column
	m := make(map[string]Column)

	// 先放默认
	for _, col := range defaultCols {
		m[col.Key] = col
	}

	// 再用自定义覆盖 / 添加
	for _, col := range customCols {
		m[col.Key] = col
	}

	// 保持顺序（默认顺序 + 新增追加）
	var result []Column

	used := make(map[string]bool)

	for _, col := range defaultCols {
		if c, ok := m[col.Key]; ok {
			result = append(result, c)
			used[col.Key] = true
		}
	}

	for _, col := range customCols {
		if !used[col.Key] {
			result = append(result, col)
		}
	}

	// 按照 Order 排序，如果 Order 相同则保持原有顺序
	sort.SliceStable(result, func(i, j int) bool {
		return result[i].Order < result[j].Order
	})

	return result
}

func ApplyColumns[T any](
	sheet *excelkit.SheetBuilder[map[string]any],
	customCols ...Column,
) {
	if sheet == nil {
		return
	}

	defaultCols := ParseColumns[T]()
	cols := MergeColumns(defaultCols, customCols)

	for _, col := range cols {
		c := col
		if c.Ignore {
			continue
		}
		if c.Key == "" && c.Value == nil {
			continue
		}

		valFn := c.Value
		if valFn == nil {
			valFn = func(m map[string]any) any {
				return m[c.Key]
			}
		}

		sheet.
			Column(c.Header).
			Value(valFn).
			End()
	}
}

func OutputMapDataForStruct[T any](data any, name string, filename string, w http.ResponseWriter, customCols ...Column) error {
	return OutputMapDataForStructWithStyles[T](data, name, filename, w, excelkit.TableHeaderBlueStyle(), excelkit.TableBodyBlueStyle(), customCols...)
}

func OutputMapDataForStructWithTitleSubtitle[T any](
	data any,
	name string,
	title string,
	subtitle string,
	filename string,
	w http.ResponseWriter,
	customCols ...Column,
) error {
	return OutputMapDataForStructWithTitleSubtitleAndStyles[T](
		data,
		name,
		title,
		subtitle,
		filename,
		w,
		nil,
		nil,
		excelkit.TableHeaderBlueStyle(),
		excelkit.TableBodyBlueStyle(),
		customCols...,
	)
}

func OutputMapDataForStructWithTitleSubtitleAndStyles[T any](
	data any,
	name string,
	title string,
	subtitle string,
	filename string,
	w http.ResponseWriter,
	titleStyle *excelkit.Style,
	subtitleStyle *excelkit.Style,
	headerStyle *excelkit.Style,
	bodyStyle *excelkit.Style,
	customCols ...Column,
) error {
	rows, ok := toMapSlice(data)
	if !ok {
		return &excelkitError{Message: "data must be a slice of map[string]any"}
	}

	builder := excelkit.New[map[string]any]().
		FromSlice(rows).
		UseStream()

	sheet := builder.Sheet(name).
		Title(title).
		Subtitle(subtitle).
		TitleStyle(titleStyle).
		SubtitleStyle(subtitleStyle).
		HeaderStyle(headerStyle).
		SheetDefaultStyle(bodyStyle)

	ApplyColumns[T](sheet, customCols...)

	return sheet.
		EndSheet().
		Download(w, filename)
}

func OutputMapDataForStructWithStyles[T any](
	data any,
	name string,
	filename string,
	w http.ResponseWriter,
	headerStyle *excelkit.Style,
	bodyStyle *excelkit.Style,
	customCols ...Column,
) error {
	rows, ok := toMapSlice(data)
	if !ok {
		return &excelkitError{Message: "data must be a slice of map[string]any"}
	}

	builder := excelkit.New[map[string]any]().
		FromSlice(rows).
		UseStream()

	sheet := builder.Sheet(name).
		HeaderStyle(headerStyle).
		SheetDefaultStyle(bodyStyle)

	ApplyColumns[T](sheet, customCols...)

	return sheet.
		EndSheet().
		Download(w, filename)
}

// EncodeRFC5987 将 UTF-8 字符串按 RFC 5987 的规则进行百分号编码
func EncodeRFC5987(s string) string {
	return url.PathEscape(s)
}

type excelkitError struct {
	Message string
}

func (e *excelkitError) Error() string { return e.Message }

func toMapSlice(data any) ([]map[string]any, bool) {
	if data == nil {
		return nil, false
	}
	if v, ok := data.([]map[string]any); ok {
		return v, true
	}

	rv := reflect.ValueOf(data)
	if rv.Kind() != reflect.Slice {
		return nil, false
	}
	out := make([]map[string]any, 0, rv.Len())
	for i := 0; i < rv.Len(); i++ {
		mv := rv.Index(i)
		if mv.Kind() == reflect.Interface {
			mv = mv.Elem()
		}
		if mv.Kind() != reflect.Map {
			return nil, false
		}
		row := map[string]any{}
		iter := mv.MapRange()
		for iter.Next() {
			k := iter.Key()
			if k.Kind() != reflect.String {
				return nil, false
			}
			row[k.String()] = iter.Value().Interface()
		}
		out = append(out, row)
	}
	return out, true
}

// ApplyColumns[VdFinanceVehicleUseUnitDetailVO](
// 	sheet,
// 	Column{
// 		Key:    "amount",
// 		Header: "金额(元)",
// 	},
// )

// ApplyColumns[VdFinanceVehicleUseUnitDetailVO](
// 	sheet,
// 	Column{
// 		Key:    "amount_fmt",
// 		Header: "金额(格式化)",
// 		Value: func(m map[string]any) any {
// 			if v, ok := m["amount"].(float64); ok {
// 				return fmt.Sprintf("%.2f 元", v)
// 			}
// 			return ""
// 		},
// 	},
// )

// idx := 0

// ApplyColumns[VdFinanceVehicleUseUnitDetailVO](
// 	sheet,
// 	Column{
// 		Key:    "_index",
// 		Header: "序号",
// 		Value: func(m map[string]any) any {
// 			idx++
// 			return idx
// 		},
// 	},
// )

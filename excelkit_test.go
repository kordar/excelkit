package excelkit

import (
	"bytes"
	"fmt"
	"math/rand"
	"net/http"
	"net/http/httptest"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// --------------------
// 模型定义
// --------------------
type User struct {
	ID        int64     `json:"id"`
	Name      string    `json:"name"`
	Email     string    `json:"email"`
	Status    string    `json:"status"`
	CreatedAt time.Time `json:"created_at"`

	// 辅助字段，用于小结
	IsSummary    bool `json:"-"`
	SummaryCount int  `json:"-"`
}

// --------------------
// TestExport 测试
// --------------------
func TestExport(t *testing.T) {
	rand.Seed(time.Now().UnixNano())

	// 测试数据
	users := make([]User, 0, 50)
	for i := 1; i <= 50; i++ {
		users = append(users, User{
			ID:        int64(i),
			Name:      fmt.Sprintf("用户%d", i),
			Email:     fmt.Sprintf("user%d@example.com", i),
			Status:    []string{"启用", "禁用"}[rand.Intn(2)],
			CreatedAt: time.Now().Add(time.Duration(-i) * time.Hour),
		})
	}

	var processedCount int

	// DSL + 高级拦截器
	// 注意：这里我们手动定义 Column 而不是 AutoSheet，以便处理小结行的特殊显示逻辑
	b := New[User]().
		FromSlice(users).
		UseStream()

	sheet := b.Sheet("用户表").
		Title("用户表").
		Subtitle("导出示例（标题居中，副标题左对齐）")
	sheet.Column("ID").Value(func(u User) any {
		if u.IsSummary {
			return "小结"
		}
		return u.ID
	}).Style(func(u User) *Style {
		if u.IsSummary {
			return nil // Let merge range style take precedence or handle separately if needed, but summary cells are merged
		}
		return BorderedStyle()
	}).End()

	sheet.Column("姓名").Value(func(u User) any {
		if u.IsSummary {
			return u.SummaryCount
		}
		return u.Name
	}).Style(func(u User) *Style { return BorderedStyle() }).End()

	sheet.Column("邮箱").Value(func(u User) any {
		return u.Email
	}).Style(func(u User) *Style { return BorderedStyle() }).End()

	sheet.Column("状态").Value(func(u User) any {
		return u.Status
	}).Style(func(u User) *Style { return BorderedStyle() }).End()

	sheet.Column("创建时间").Value(func(u User) any {
		return u.CreatedAt
	}).Style(func(u User) *Style { return BorderedStyle() }).End()

	sheet.EndSheet()

	b.AddInterceptor(func(row User, rowIndex int) (bool, []User, []MergeRange[User]) {
		processedCount++

		extraRows := []User{}
		mergeRanges := []MergeRange[User]{}

		// 每 10 行插入小结行
		// 注意：rowIndex 是 Excel 中的行号（包含表头）
		// 表头占 1 行。数据从第 2 行开始。
		// 当处理完第 10 条数据时，rowIndex 指向的是第 12 行（因为 interceptor 是在写入当前 row 之前调用的？）
		// Wait, Interceptor is called BEFORE writing `row`.
		// If we are about to write the 11th data row (User 11), rowIndex is 12 (1 header + 10 data).
		// We want to insert summary AFTER User 10.
		// So when we see User 11 (Row 12), we insert summary.
		// (12-1) % 10 != 0. 11 % 10 = 1.
		// Wait, previous logic: `if rowIndex > 1 && (rowIndex-1)%10 == 0`.
		// Row 11: (11-1)%10 = 0. This is User 10.
		// So BEFORE writing User 10, we insert summary? That puts summary between User 9 and User 10.
		// We want summary AFTER User 10.
		// So we should insert it when we are at User 11 (Row 12)?
		// OR, we insert it *after* User 10 in the stream?
		// The interceptor returns `extraRows` which are written *before* the current row.
		// So if we want summary after User 10, we must insert it when processing User 11.
		// User 11 is at Row 12.
		// So check should be `(rowIndex-2) % 10 == 0`? (If row 12, 10%10==0).
		// Let's adjust logic.
		//
		// Actually, let's look at `processedCount`.
		// If `processedCount` is 11 (User 11). We want summary after 10.
		// So if `(processedCount - 1) % 10 == 0`?
		// When processedCount is 1, no.
		// When processedCount is 11. We want summary.

		if processedCount > 1 && (processedCount-1)%10 == 0 {
			summary1 := User{
				IsSummary:    true,
				SummaryCount: processedCount - 1, // Summary of previous 10
			}
			summary2 := User{IsSummary: true}
			extraRows = append(extraRows, summary1, summary2)

			// 当前 rowIndex 是即将写入 User 11 的位置。
			// 我们插入了 extraRows (summary)，它会占用 rowIndex。
			// User 11 会被推到 rowIndex + 1。
			// MergeRange 需要指定 summary 所在的行。
			// summary 将被写入当前的 rowIndex。

			mergeRanges = append(mergeRanges, MergeRange[User]{
				StartRow: rowIndex,
				EndRow:   rowIndex,
				StartCol: 2,
				EndCol:   5,
				RowData:  summary1,
				Style:    BlueSummaryBorderedStyle(),
			})

			mergeRanges = append(mergeRanges, MergeRange[User]{
				StartRow: rowIndex + 1,
				EndRow:   rowIndex + 1,
				StartCol: 2,
				EndCol:   5,
				RowData:  summary2,
				Style:    BorderedStyle(),
			})
		}

		return true, extraRows, mergeRanges
	})

	err := b.Save("users_final_advanced.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	f, err := excelize.OpenFile("users_final_advanced.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	defer func() {
		_ = f.Close()
	}()

	mergeCells, err := f.GetMergeCells("用户表", true)
	if err != nil {
		t.Fatal(err)
	}

	summaryCount := (len(users) - 1) / 10
	expected := map[string]bool{}
	for k := 1; k <= summaryCount; k++ {
		startRow := 12*k + 2
		expected[fmt.Sprintf("B%d:E%d", startRow, startRow)] = false
		expected[fmt.Sprintf("B%d:E%d", startRow+1, startRow+1)] = false
	}
	for _, mc := range mergeCells {
		rng := mc.GetStartAxis() + ":" + mc.GetEndAxis()
		if _, ok := expected[rng]; ok {
			expected[rng] = true
		}
	}
	for rng, ok := range expected {
		if !ok {
			t.Fatalf("missing merged range: %s", rng)
		}
	}

	name11, err := f.GetCellValue("用户表", "B16")
	if err != nil {
		t.Fatal(err)
	}
	if name11 != "用户11" {
		t.Fatalf("unexpected B16 value: %q", name11)
	}

	t.Log("导出成功: users_final_advanced.xlsx")
}

func TestDownload(t *testing.T) {
	users := []User{{ID: 1, Name: "Test"}}
	b := New[User]().FromSlice(users).AutoSheet("json", "Sheet1").EndSheet()

	// Test Write
	buf := &bytes.Buffer{}
	if err := b.Write(buf); err != nil {
		t.Fatal(err)
	}
	if buf.Len() == 0 {
		t.Fatal("Buffer is empty")
	}

	// Test Download
	rec := httptest.NewRecorder()
	if err := b.Download(rec, "test.xlsx"); err != nil {
		t.Fatal(err)
	}
	if rec.Header().Get("Content-Disposition") != "attachment; filename=\"test.xlsx\"; filename*=UTF-8''test.xlsx" {
		t.Errorf("Unexpected Content-Disposition: %s", rec.Header().Get("Content-Disposition"))
	}
	if rec.Body.Len() == 0 {
		t.Fatal("Response body is empty")
	}

	// Test Download with Custom Headers
	rec2 := httptest.NewRecorder()
	customHeaders := http.Header{}
	customHeaders.Set("X-Custom-Header", "ExcelKit-Test")
	customHeaders.Set("Content-Type", "application/vnd.ms-excel")

	if err := b.Download(rec2, "test_custom.xlsx", customHeaders); err != nil {
		t.Fatal(err)
	}

	if rec2.Header().Get("Content-Disposition") != "attachment; filename=\"test_custom.xlsx\"; filename*=UTF-8''test_custom.xlsx" {
		t.Errorf("Unexpected Content-Disposition: %s", rec2.Header().Get("Content-Disposition"))
	}
	if rec2.Header().Get("X-Custom-Header") != "ExcelKit-Test" {
		t.Errorf("Custom header missing")
	}
	if rec2.Header().Get("Content-Type") != "application/octet-stream" {
		t.Errorf("Unexpected Content-Type: got %s", rec2.Header().Get("Content-Type"))
	}
}

func TestCustomSheetNameNoDefaultSheet1(t *testing.T) {
	rows := []User{{ID: 1, Name: "A"}}

	buf := &bytes.Buffer{}
	if err := New[User]().
		FromSlice(rows).
		UseStream().
		AutoSheet("json", "用户表").
		EndSheet().
		Write(buf); err != nil {
		t.Fatal(err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	for _, name := range f.GetSheetList() {
		if name == "Sheet1" {
			t.Fatalf("unexpected default sheet: %s", name)
		}
	}
}

func TestHeaderAndDefaultStyle(t *testing.T) {
	rows := []User{{ID: 1, Name: "A"}}

	buf := &bytes.Buffer{}
	b := New[User]().FromSlice(rows).UseStream()
	b.Sheet("S").
		HeaderStyle(BorderedHeaderStyle()).
		SheetDefaultStyle(BorderedStyle()).
		Column("ID").Value(func(u User) any { return u.ID }).End().
		Column("姓名").Value(func(u User) any { return u.Name }).Style(func(u User) *Style { return GreenStyle() }).End().
		EndSheet()

	if err := b.Write(buf); err != nil {
		t.Fatal(err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	hA1, err := f.GetCellStyle("S", "A1")
	if err != nil {
		t.Fatal(err)
	}
	hB1, err := f.GetCellStyle("S", "B1")
	if err != nil {
		t.Fatal(err)
	}
	dA2, err := f.GetCellStyle("S", "A2")
	if err != nil {
		t.Fatal(err)
	}
	dB2, err := f.GetCellStyle("S", "B2")
	if err != nil {
		t.Fatal(err)
	}

	if hA1 == 0 || hA1 != hB1 {
		t.Fatalf("unexpected header style ids: A1=%d B1=%d", hA1, hB1)
	}
	if dA2 == 0 {
		t.Fatalf("unexpected default style id: A2=%d", dA2)
	}
	if dA2 == hA1 {
		t.Fatalf("expected data style differs from header style: data=%d header=%d", dA2, hA1)
	}
	if dB2 == dA2 {
		t.Fatalf("expected column override style differs from default: A2=%d B2=%d", dA2, dB2)
	}
}

func TestTitleSubtitleStyle(t *testing.T) {
	rows := []User{{ID: 1, Name: "A"}}

	buf := &bytes.Buffer{}
	b := New[User]().FromSlice(rows).UseStream()
	b.Sheet("S").
		Title("T").
		TitleStyle(TableHeaderDarkStyle()).
		Subtitle("Sub").
		SubtitleStyle(TableHeaderLightStyle()).
		HeaderStyle(BorderedHeaderStyle()).
		Column("ID").Value(func(u User) any { return u.ID }).End().
		Column("姓名").Value(func(u User) any { return u.Name }).End().
		EndSheet()

	if err := b.Write(buf); err != nil {
		t.Fatal(err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	tA1, err := f.GetCellStyle("S", "A1")
	if err != nil {
		t.Fatal(err)
	}
	tB1, err := f.GetCellStyle("S", "B1")
	if err != nil {
		t.Fatal(err)
	}
	sA2, err := f.GetCellStyle("S", "A2")
	if err != nil {
		t.Fatal(err)
	}
	sB2, err := f.GetCellStyle("S", "B2")
	if err != nil {
		t.Fatal(err)
	}
	hA3, err := f.GetCellStyle("S", "A3")
	if err != nil {
		t.Fatal(err)
	}

	if tA1 == 0 || tA1 != tB1 {
		t.Fatalf("unexpected title style ids: A1=%d B1=%d", tA1, tB1)
	}
	if sA2 == 0 || sA2 != sB2 {
		t.Fatalf("unexpected subtitle style ids: A2=%d B2=%d", sA2, sB2)
	}
	if tA1 == sA2 {
		t.Fatalf("expected different style ids for title/subtitle: title=%d subtitle=%d", tA1, sA2)
	}
	if hA3 == 0 {
		t.Fatalf("unexpected header style id: %d", hA3)
	}
	if hA3 == tA1 || hA3 == sA2 {
		t.Fatalf("expected header style differs from title/subtitle: header=%d title=%d subtitle=%d", hA3, tA1, sA2)
	}
}

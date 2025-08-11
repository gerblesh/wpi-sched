package cmd

import (
	"fmt"
	"io"
	"os"
	"regexp"
	"slices"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var (
	rootCmd = &cobra.Command{
		Use:   "wpi-sched",
		Short: "wpi-sched exports your *incredible* xsls schedule from workday into an actually usable calendar format (*.ics) :)",
		RunE:  Export,
	}
	fExcelSheet string
	fOutput     string
)

const (
	MEETING_COL    string = "Meeting Patterns"
	DESC_COL       string = "Course Listing"
	START_DATE_COL string = "Start Date"
	END_DATE_COL   string = "End Date"
	INSTRUCTOR_COL string = "Instructor"
	TIMEZONE       string = "America/New_York"
)

type Course struct {
	Description string
	Meeting     string
	StartDate   string
	EndDate     string
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func init() {
	rootCmd.PersistentFlags().StringVarP(&fExcelSheet, "file", "f", "View_My_Courses.xlsx", "Excel file containing schedule info")
	rootCmd.PersistentFlags().StringVarP(&fOutput, "output", "o", "", "ics output containing schedule info")
}

func Export(cmd *cobra.Command, args []string) error {
	f, err := excelize.OpenFile(fExcelSheet)
	if err != nil {
		return err
	}
	defer func() error {
		return f.Close()
	}()
	courses, err := GetCourses(f)
	if err != nil {
		return err
	}

	w, err := getWriter(fOutput)
	if err != nil {
		return err
	}

	err = WriteIcalBuf(courses, w)
	return err
}

func getWriter(path string) (io.Writer, error) {
	if path == "" || path == "-" {
		return os.Stdout, nil
	}
	f, err := os.Create(path)
	if err != nil {
		return nil, err
	}
	return f, nil
}

func getColumns(rows [][]string) (map[string]int, int, error) {
	cols := map[string]int{
		MEETING_COL:    -1,
		DESC_COL:       -1,
		START_DATE_COL: -1,
		END_DATE_COL:   -1,
	}
	startRow := -1
	// Find column indices
	for i, row := range rows {
		for j, colCell := range row {
			switch colCell {
			case DESC_COL:
				cols[DESC_COL] = j
				startRow = i + 1
			case MEETING_COL:
				cols[MEETING_COL] = j
				startRow = i + 1
			case START_DATE_COL:
				cols[START_DATE_COL] = j
				startRow = i + 1
			case END_DATE_COL:
				cols[END_DATE_COL] = j
				startRow = i + 1
			case INSTRUCTOR_COL:
				cols[INSTRUCTOR_COL] = j
				startRow = i + 1
			}
		}
		if startRow != -1 {
			break
		}

	}

	// check for columns that weren't found
	for k, v := range cols {
		if v == -1 {
			return cols, startRow, fmt.Errorf("Unable to Find Column: %v", k)
		}
	}
	return cols, startRow, nil
}

func GetCourses(f *excelize.File) ([]Course, error) {

	courses := []Course{}
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return courses, err
	}

	cols, startRow, err := getColumns(rows)
	if err != nil {
		return courses, fmt.Errorf("Unabled to parse Columns: %v", err)
	}
	startDateCol := cols[START_DATE_COL]
	endDateCol := cols[END_DATE_COL]
	meetingCol := cols[MEETING_COL]
	descriptionCol := cols[DESC_COL]
	instructorCol := cols[INSTRUCTOR_COL]

	for _, row := range rows[startRow:] {
		if len(row) <= endDateCol {
			break
		}
		descSlice := []string{}
		descSlice = append(descSlice, row[descriptionCol])
		if row[instructorCol] != "" {
			descSlice = append(descSlice, row[instructorCol])
		}
		courses = append(courses, Course{
			strings.Join(descSlice, " - "),
			row[meetingCol],
			row[startDateCol],
			row[endDateCol],
		})
	}
	return courses, nil
}

func WriteIcalBuf(courses []Course, w io.Writer) error {
	_, err := fmt.Fprintf(w, "BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n")
	if err != nil {
		return err
	}
	for _, c := range courses {
		e, err := GetIcal(c)
		if err != nil {
			return err
		}
		_, err = fmt.Fprint(w, e)
		if err != nil {
			return err
		}
	}
	_, err = fmt.Fprintf(w, "END:VCALENDAR\n")
	return err
}

/*
Parses raw input from the table and turns it into an ical event
Input:
desc: simple course description (unchanged)
meeting: meeting patterns, comes in formatted roughly like: M-T-W-R-F | HH:MM AM - HH:MM AM | LOCATION

The rest is pretty self explanitory
endDate: MM/DD/YYYY
startDate: MM/DD/YYYY
*/
func GetIcal(c Course) (string, error) {
	// rough reference: https://gist.github.com/DeMarko/6142417
	// actual spec: https://www.rfc-editor.org/rfc/rfc5545
	type day struct {
		ICal string
		Time time.Weekday
	}
	dayMap := map[string]day{
		"M": {"MO", time.Monday},
		"T": {"TU", time.Tuesday},
		"W": {"WE", time.Wednesday},
		"R": {"TH", time.Thursday},
		"F": {"FR", time.Friday},
	}
	// Parsing the 'Meeting Patterns' column
	// sometimes the 'Meeting Patterns' column can just say fuck all for specific classes, literally nothing useful here, just give up
	if c.Meeting == "" {
		return "", nil
	}
	parsedMeet := strings.Split(c.Meeting, "|")
	if len(parsedMeet) < 3 {
		return "", fmt.Errorf("unable to parse 'Meeting Patterns': expected at least 3 parts, got %d", len(parsedMeet))
	}

	freq := strings.TrimSpace(parsedMeet[0])
	parsedFreq := []string{}
	validWeekdays := []time.Weekday{}
	for d := range strings.SplitSeq(freq, "-") {
		day, ok := dayMap[d]
		if !ok {
			return "", fmt.Errorf("unable to parse day: %s, not in map: %v", d, dayMap)
		}
		parsedFreq = append(parsedFreq, day.ICal)
		validWeekdays = append(validWeekdays, day.Time)
	}
	byDay := strings.Join(parsedFreq, ",")

	endDate, err := convertDate(c.EndDate)
	if err != nil {
		return "", err
	}
	startDate, err := convertStartDate(c.StartDate, validWeekdays)
	if err != nil {
		return "", err
	}

	times := strings.Split(strings.TrimSpace(parsedMeet[1]), "-")
	if len(times) < 2 {
		return "", fmt.Errorf("unable to parse 'Meeting Patterns': expected at least 2 times (start and end), got %d", len(parsedMeet))
	}
	startTime, err := convertTime(strings.TrimSpace(times[0]))
	if err != nil {
		return "", err
	}
	endTime, err := convertTime(strings.TrimSpace(times[1]))
	if err != nil {
		return "", err
	}

	location := strings.TrimSpace(parsedMeet[2])

	ical := fmt.Sprintf(`BEGIN:VEVENT
UID:%s
DTSTAMP:%s
DTSTART;TZID=%s:%sT%s
DTEND;TZID=%s:%sT%s
SUMMARY:%s
LOCATION:%s
RRULE:FREQ=WEEKLY;BYDAY=%s;UNTIL=%sT235959
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Reminder - %s starts soon
END:VALARM
END:VEVENT
`,
		sanitizeUID(c.Description, byDay),
		time.Now().UTC().Format("20060102T150405Z"),
		TIMEZONE, startDate, startTime,
		TIMEZONE, startDate, endTime,
		c.Description,
		location,
		byDay,
		endDate,
		c.Description,
	)
	return ical, nil
}

func sanitizeUID(desc string, days string) string {
	re := regexp.MustCompile(`[^a-zA-Z0-9]+`)
	cleaned := re.ReplaceAllString(desc+days, "_")
	return strings.Trim(cleaned, "_") + "@wpi.edu"
}

// convertDate converts "MM/DD/YYYY" -> "YYYYMMDD" and moves the date to a valid weekday to make ical happy
func convertStartDate(dateStr string, validDays []time.Weekday) (string, error) {
	t, err := time.Parse("01-02-06", dateStr)
	if err != nil {
		return "", err
	}
	for range 7 {
		if slices.Contains(validDays, t.Weekday()) {
			break
		}
		t = t.AddDate(0, 0, 1)
	}
	return t.Format("20060102"), nil
}

// convertDate converts "MM-DD-YYYY" -> "YYYYMMDD" and moves the date to a valid weekday to make ical happy
func convertDate(dateStr string) (string, error) {
	t, err := time.Parse("01-02-06", dateStr)
	if err != nil {
		return "", err
	}
	return t.Format("20060102"), nil
}

// convertTime converts "HH:MM AM/PM" -> "HHMMSS" (24-hour format)
func convertTime(timeStr string) (string, error) {
	t, err := time.Parse("3:04 PM", timeStr)
	if err != nil {
		return "", err
	}
	return t.Format("150405"), nil
}

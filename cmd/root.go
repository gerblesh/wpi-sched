package cmd

import (
	"fmt"
	"io"
	"os"
	"regexp"
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

	for _, row := range rows[startRow:] {
		if len(row) <= endDateCol {
			break
		}
		courses = append(courses, Course{
			row[descriptionCol],
			row[meetingCol],
			row[startDateCol],
			row[endDateCol],
		})
	}
	return courses, nil
}

func WriteIcalBuf(courses []Course, w io.Writer) error {
	_, err := w.Write([]byte("BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN"))
	if err != nil {
		return err
	}
	for _, c := range courses {
		e, err := GetIcal(c)
		if err != nil {
			return err
		}
		_, err = w.Write(e)
		if err != nil {
			return err
		}
	}
	_, err = w.Write([]byte("\nEND:VCALENDAR\n"))
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
func GetIcal(c Course) ([]byte, error) {
	// rough reference: https://gist.github.com/DeMarko/6142417
	// actual spec: https://www.rfc-editor.org/rfc/rfc5545
	dayMap := map[string]string{
		"M": "MO",
		"T": "TU",
		"W": "WE",
		"R": "TH",
		"F": "FR",
	}

	// Parsing the 'Meeting Patterns' column
	parsedMeet := strings.Split(c.Meeting, "|")
	if len(parsedMeet) < 3 {
		return nil, fmt.Errorf("unable to parse 'Meeting Patterns': expected at least 3 parts, got %d", len(parsedMeet))
	}

	freq := strings.TrimSpace(parsedMeet[0])
	parsedFreq := []string{}
	for d := range strings.SplitSeq(freq, "-") {
		day, ok := dayMap[d]
		if !ok {
			return nil, fmt.Errorf("unable to parse day: %s, not in map: %v", d, dayMap)
		}
		parsedFreq = append(parsedFreq, day)
	}
	byDay := strings.Join(parsedFreq, ",")

	times := strings.Split(strings.TrimSpace(parsedMeet[1]), "-")
	if len(times) < 2 {
		return nil, fmt.Errorf("unable to parse 'Meeting Patterns': expected at least 2 times (start and end), got %d", len(parsedMeet))
	}
	startTime, err := convertTime(strings.TrimSpace(times[0]))
	if err != nil {
		return nil, err
	}
	endTime, err := convertTime(strings.TrimSpace(times[0]))
	if err != nil {
		return nil, err
	}

	endDateStr, err := convertDate(c.EndDate)
	if err != nil {
		return nil, err
	}
	startDateStr, err := convertDate(c.StartDate)
	if err != nil {
		return nil, err
	}

	location := strings.TrimSpace(parsedMeet[2])

	ical := fmt.Sprintf(`BEGIN:VEVENT
UID:%s
DTSTAMP:%s
DTSTART;TZID=America/New_York:%sT%s
DTEND;TZID=America/New_York:%sT%s
SUMMARY:%s
LOCATION:%s
RRULE:FREQ=WEEKLY;BYDAY=%s;UNTIL=%sT235959
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Reminder - %s starts soon
END:VALARM
END:VEVENT`,
		sanitizeUID(c.Description),
		time.Now().UTC().Format("20060102T150405Z"),
		startDateStr, startTime,
		startDateStr, endTime,
		c.Description,
		location,
		byDay,
		endDateStr,
		c.Description,
	)
	return []byte(ical), nil
}

func sanitizeUID(desc string) string {
	re := regexp.MustCompile(`[^a-zA-Z0-9]+`)
	cleaned := re.ReplaceAllString(desc, "_")
	return strings.Trim(cleaned, "_") + "@wpi.edu"
}

// convertDate converts "MM/DD/YYYY" -> "YYYYMMDD"
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

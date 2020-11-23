package main

import (
	"context"
	"fmt"
	"io/ioutil"
	"os"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

type SheetClient struct {
	srv           *sheets.Service
	spreadsheetID string
}

func NewSheetClient(ctx context.Context, spreadsheetID string) (*SheetClient, error) {
	b, err := ioutil.ReadFile("secret.json")
	if err != nil {
		return nil, err
	}
	jwt, err := google.JWTConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		return nil, err
	}
	srv, err := sheets.New(jwt.Client(ctx))
	if err != nil {
		return nil, err
	}
	return &SheetClient{
		srv:           srv,
		spreadsheetID: spreadsheetID,
	}, nil
}

// https://developers.google.com/sheets/api/guides/values#reading_a_single_range
func (s *SheetClient) Get(range_ string) ([][]interface{}, error) {
	resp, err := s.srv.Spreadsheets.Values.Get(s.spreadsheetID, range_).Do()
	if err != nil {
		return nil, err
	}
	return resp.Values, nil
}

// https://developers.google.com/sheets/api/guides/values#writing_to_a_single_range
func (s *SheetClient) Update(range_ string, values [][]interface{}) error {
	_, err := s.srv.Spreadsheets.Values.Update(s.spreadsheetID, range_, &sheets.ValueRange{
		Values: values,
	}).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		return err
	}
	return nil
}

// https://developers.google.com/sheets/api/samples/formatting
func (s *SheetClient) Format(range_ *sheets.GridRange, format *sheets.CellFormat) error {
	_, err := s.srv.Spreadsheets.BatchUpdate(s.spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				RepeatCell: &sheets.RepeatCellRequest{
					Fields: "userEnteredFormat(backgroundColor)",
					Range:  range_,
					Cell: &sheets.CellData{
						UserEnteredFormat: format,
					},
				},
			},
		},
	}).Do()
	if err != nil {
		return err
	}
	return nil
}

func (s *SheetClient) SheetID(sheetName string) (int64, error) {
	resp, err := s.srv.Spreadsheets.Get(s.spreadsheetID).Do()
	if err != nil {
		return 0, err
	}
	for _, sheet := range resp.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}
	return 0, fmt.Errorf("sheetName %s is not found", sheetName)
}

// https://developers.google.com/sheets/api/guides/values#appending_values
func (s *SheetClient) Append(values [][]interface{}) error {
	_, err := s.srv.Spreadsheets.Values.Append(s.spreadsheetID, "シート1", &sheets.ValueRange{
		Values: values,
	}).ValueInputOption("USER_ENTERED").InsertDataOption("INSERT_ROWS").Do()
	if err != nil {
		return err
	}
	return nil
}

func (s *SheetClient) List(range_ *sheets.GridRange, values []string) error {
	conditionValues := make([]*sheets.ConditionValue, 0, len(values))
	for _, value := range values {
		conditionValues = append(conditionValues, &sheets.ConditionValue{
			UserEnteredValue: value,
		})
	}
	_, err := s.srv.Spreadsheets.BatchUpdate(s.spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				RepeatCell: &sheets.RepeatCellRequest{
					Fields: "dataValidation",
					Range:  range_,
					Cell: &sheets.CellData{
						DataValidation: &sheets.DataValidationRule{
							Condition: &sheets.BooleanCondition{
								Type:   "ONE_OF_LIST",
								Values: conditionValues,
							},
							ShowCustomUi: true,
							Strict:       true,
						},
					},
				},
			},
		},
	}).Do()
	if err != nil {
		return err
	}
	return nil
}

func main() {
	ctx := context.Background()
	client, err := NewSheetClient(ctx, os.Getenv("SPREAD_SHEET_ID"))
	if err != nil {
		panic(err)
	}
	sheetID, err := client.SheetID("シート1")
	if err != nil {
		panic(err)
	}

	if err := client.Update("A1", [][]interface{}{
		{
			"aaa",
			"bbb",
		},
		{
			"ccc",
			"ddd",
		},
	}); err != nil {
		panic(err)
	}

	if err := client.Append([][]interface{}{
		{
			"1",
		},
	}); err != nil {
		panic(err)
	}

	if err := client.Format(&sheets.GridRange{
		SheetId:          sheetID,
		StartColumnIndex: 1,
		StartRowIndex:    2,
		EndColumnIndex:   3,
		EndRowIndex:      4,
	}, &sheets.CellFormat{
		BackgroundColor: &sheets.Color{
			Red: 1.0,
		},
	}); err != nil {
		panic(err)
	}

	values, err := client.Get("'シート1'!A1:B1")
	for _, row := range values {
		fmt.Println(row)
	}

	if err := client.List(&sheets.GridRange{
		SheetId:          sheetID,
		StartColumnIndex: 1,
		StartRowIndex:    2,
		EndColumnIndex:   3,
		EndRowIndex:      4,
	}, []string{"○", "×"}); err != nil {
		panic(err)
	}

}

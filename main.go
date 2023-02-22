package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"reflect"
	"strconv"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

const spreadsheetId = "1FLfk_jy3Hao0hBem2jtCdOBPtYXeYRxBJ51tahSanPQ"
const sheetName = "Users"

func main() {
	// Initialize credentials
	b, err := os.ReadFile("credentials.json") // https://developers.google.com/identity/protocols/oauth2/service-account
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// Initialize connection
	config, err := google.JWTConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	srv, err := sheets.NewService(context.Background(), option.WithHTTPClient(config.Client(context.Background())))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	// Make sure there is a header row and it matches what is expected
	validateHeaders(srv)

	// Demonstrate writing, updating, and adding rows
	row := RowEntry{"John", "1/1/1990", "100", "Gray"}
	write(srv, sheetName+"!A5", row.stoi())
	update(srv, "John", []interface{}{"Jane"})
	add(srv, row.stoi())
}

func update(srv *sheets.Service, pk string, values []interface{}) {
	writeRange := sheetName + "!A"
	resp := query(srv, sheetName+"!A2:D")
	if len(resp) == 0 {
		fmt.Println("No data found.")
	} else {
		for i, row := range resp {
			if len(row) != 4 {
				continue
			}
			var entry RowEntry
			entry.itos(row)
			// TODO: Change to reflect field by name
			if entry.Name == pk {
				fmt.Println("(" + pk + ") found on cell A" + strconv.Itoa(i+2))
				writeRange += strconv.Itoa(i + 2)
				fmt.Println("Write Range: " + writeRange)
				write(srv, writeRange, values)
				break
			}
		}
	}
}

func add(srv *sheets.Service, values []interface{}) {
	nr := len(query(srv, sheetName+"!A:B")) + 1
	endRange := sheetName + "!A" + fmt.Sprint(nr)
	write(srv, endRange, values)
}

func query(srv *sheets.Service, rng string) [][]interface{} {
	resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, rng).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}
	return resp.Values
}

func write(srv *sheets.Service, rng string, values []interface{}) {
	var vr sheets.ValueRange
	vr.Values = append(vr.Values, values)
	_, err := srv.Spreadsheets.Values.Update(spreadsheetId, rng, &vr).ValueInputOption("RAW").Do()
	if err != nil {
		log.Fatalf("Unable to write to sheet: %v", err)
	}
}

func delete(srv *sheets.Service, rng string) {
	var cvr sheets.ClearValuesRequest
	_, err := srv.Spreadsheets.Values.Clear(spreadsheetId, rng, &cvr).Do()
	if err != nil {
		fmt.Println("Error clearing header row")
	}
}

func validateHeaders(srv *sheets.Service) {
	headerRow := sheetName + "!A1:D1"
	defHeader := []interface{}{}
	refVal := reflect.ValueOf(RowEntry{})
	for i := 0; i < refVal.NumField(); i++ {
		defHeader = append(defHeader, refVal.Type().Field(i).Name)
	}
	resp := query(srv, headerRow)
	if len(resp) == 0 {
		fmt.Println("No header found. Adding...")
		write(srv, headerRow, defHeader)
	} else {
		currHeader := resp[0]
		if reflect.DeepEqual(defHeader, currHeader) != true {
			fmt.Printf("Defined Header: %+v\n", defHeader)
			fmt.Printf("Current Header: %+v\n", currHeader)
			log.Fatal("Mismatched headers.")
		}
	}
}

// Correspond to header row of the spreadsheet
type RowEntry struct {
	Name      string
	Birthdate string
	Age       string
	Eyecolor  string
}

// Convert struct to []interface{}
func (u RowEntry) stoi() []interface{} {
	refVal := reflect.ValueOf(u)
	val := make([]interface{}, refVal.NumField())
	for i := 0; i < refVal.NumField(); i++ {
		val[i] = refVal.Field(i).Interface()
	}
	return val
}

// Convert response row into User struct
func (u *RowEntry) itos(r []interface{}) {
	u.Name = r[0].(string)
	u.Birthdate = r[1].(string)
	u.Age = r[2].(string)
	u.Eyecolor = r[3].(string)
}

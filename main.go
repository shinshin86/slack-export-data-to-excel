package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"golang.org/x/exp/slices"

	"github.com/kyokomi/emoji/v2"
	"github.com/xuri/excelize/v2"
)

type User struct {
	Id   string `json:"id"`
	Name string `json:"real_name"`
}

type Channel struct {
	Id   string `json:"id"`
	Name string `json:"name"`
}

type Reaction struct {
	UserIdList []string `json:"users"`
	Name       string   `json:"name"`
	Count      int      `json:"count"`
	Users      []User
}

type Replie struct {
	UserId    string `json:"user"`
	TimeStamp string `json:"ts"`
}

type Post struct {
	UserId          string `json:"user"`
	Text            string `json:"text"`
	Type            string `json:"type"`
	SubType         string `json:"subtype"`
	TimeStamp       string `json:"ts"`
	BotId           string `json:"bot_id"`
	ThreadTimeStamp string `json:"thread_ts"`
	ParentUserId    string `json:"parent_user_id"`
	User            User
	Reactions       []Reaction
	Replies         []Replie `json:"replies"`
}

func getUsers(dirname string) []User {
	raw, err := ioutil.ReadFile(filepath.Join(dirname, "users.json"))
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	var users []User

	json.Unmarshal(raw, &users)

	return users
}

func getChannels(dirname string) []Channel {
	raw, err := ioutil.ReadFile(filepath.Join(dirname, "channels.json"))
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	var channels []Channel

	json.Unmarshal(raw, &channels)

	return channels
}

func getUserNameList(users []User, f func(User) string) []string {
	r := make([]string, len(users))
	for i, u := range users {
		r[i] = f(u)
	}

	return r
}

func writeSheets(f *excelize.File, dirname string, channels []Channel, users []User) {
	// write row of cell
	var row string

	for _, channel := range channels {
		f.NewSheet(channel.Name)

		files, err := filepath.Glob(filepath.Join(dirname, channel.Name, "*.json"))
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}

		// header
		header := &[]interface{}{"index", "user", "text", "thread", "reactions", "datetime"}
		f.SetSheetRow(channel.Name, "A1", header)

		var index = 1

		for _, file := range files {
			raw, err := ioutil.ReadFile(file)

			if err != nil {
				fmt.Fprintln(os.Stderr, err)
				os.Exit(1)
			}

			var posts []Post

			json.Unmarshal(raw, &posts)

			for _, post := range posts {
				// get user
				userIndex := slices.IndexFunc(users, func(user User) bool { return user.Id == post.UserId })

				if userIndex != -1 {
					post.User = users[userIndex]
				} else if post.UserId == "USLACKBOT" {
					post.User = User{
						Id:   post.BotId,
						Name: post.UserId,
					}
				} else if post.SubType == "bot_message" {
					post.User = User{
						Id:   post.BotId,
						Name: "Bot",
					}
				}

				// convert datetime
				datetime, err := strconv.ParseInt(strings.Split(post.TimeStamp, ".")[0], 10, 64)
				if err != nil {
					fmt.Println(err)
					os.Exit(1)
				}
				dt := time.Unix(datetime, 0)

				var threadData string

				if len(post.Replies) > 0 {
					threadData = "Thread data:\n"

					for i, r := range post.Replies {
						// convert datetime
						replieDatetime, err := strconv.ParseInt(strings.Split(r.TimeStamp, ".")[0], 10, 64)
						if err != nil {
							fmt.Println(err)
							os.Exit(1)
						}
						formattedDt := time.Unix(replieDatetime, 0).Format("2006/1/2 15:04:05")

						idx := slices.IndexFunc(users, func(user User) bool { return user.Id == r.UserId })

						d := strconv.Itoa(i+1) + ": " + users[idx].Name + "(" + formattedDt + ")"

						if i < len(post.Replies)-1 {
							d = d + "\n"
						}

						threadData = threadData + d
					}
				} else if post.ParentUserId != "" {
					idx := slices.IndexFunc(users, func(user User) bool { return user.Id == post.ParentUserId })
					threadData = "Thread parent user:\n" + users[idx].Name
				}

				// excluding header (A2, A3...)
				row = "A" + strconv.Itoa(index+1)
				username := post.User.Name + " (" + post.User.Id + ")"
				f.SetSheetRow(channel.Name, row, &[]interface{}{index, username, emoji.Sprint(post.Text), threadData, "", dt})

				index++

				for _, r := range post.Reactions {
					// add Reaction.Users
					for _, id := range r.UserIdList {
						idx := slices.IndexFunc(users, func(user User) bool { return user.Id == id })
						r.Users = append(r.Users, users[idx])
					}

					reactionUsers := getUserNameList(r.Users, func(u User) string { return u.Name })

					emojiText := ":" + r.Name + ":"
					reactionsText := emoji.Sprint(emojiText + "(" + strconv.Itoa(r.Count) + ") - [" + strings.Join(reactionUsers, ",") + "]")

					// excluding header (A2, A3...)
					row = "A" + strconv.Itoa(index+1)
					f.SetSheetRow(channel.Name, row, &[]interface{}{index, "", "", "", reactionsText, dt})

					index++
				}
			}
		}
	}
}

func main() {
	flag.Parse()

	if len(flag.Args()) == 0 {
		fmt.Println("ERROR: The export data directory must be specified.")
		os.Exit(1)
	}

	exportDirName := flag.Args()[0]
	fmt.Println("Specified directory: " + exportDirName)

	users := getUsers(exportDirName)
	channels := getChannels(exportDirName)

	// Excel
	f := excelize.NewFile()
	writeSheets(f, exportDirName, channels, users)

	// Delete Sheet1(default sheet)
	f.DeleteSheet("Sheet1")

	fileName := exportDirName + ".xlsx"
	if err := f.SaveAs(fileName); err != nil {
		fmt.Println(err)
	}

	fmt.Println("SUCCESS")
}

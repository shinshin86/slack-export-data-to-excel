package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"sort"
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

type Reply struct {
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
	Replies         []Reply `json:"replies"`
}

type ThreadPost struct {
	ParentIndex    int
	ReplyUser      User
	ReplyTimeStamp string
	ParentPost     Post
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

func getRow(index int) string {
	// excluding header (A2, A3...)
	return "A" + strconv.Itoa(index+1)
}

func getUnixTime(ts string) time.Time {
	datetime, err := strconv.ParseInt(strings.Split(ts, ".")[0], 10, 64)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	return time.Unix(datetime, 0)
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

		sort.Strings(files)

		// header
		header := &[]interface{}{"index", "user", "text", "thread", "reactions", "datetime"}
		f.SetSheetRow(channel.Name, "A1", header)

		var index = 1
		var threadPosts []ThreadPost

		for _, file := range files {
			raw, err := ioutil.ReadFile(file)

			if err != nil {
				fmt.Fprintln(os.Stderr, err)
				os.Exit(1)
			}

			var posts []Post
			var reactionsText string

			json.Unmarshal(raw, &posts)

			for _, post := range posts {
				// init
				reactionsText = ""

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
				dt := getUnixTime(post.TimeStamp)

				var threadData string

				if len(post.Replies) > 0 {
					threadData = "Thread posts:\n"

					for i, r := range post.Replies {
						// convert datetime
						formattedDt := getUnixTime(r.TimeStamp).Format("2006/1/2 15:04:05")

						idx := slices.IndexFunc(users, func(user User) bool { return user.Id == r.UserId })

						d := strconv.Itoa(i+1) + ": " + users[idx].Name + "(" + formattedDt + ")"

						threadPosts = append(threadPosts, ThreadPost{
							ParentIndex:    index,
							ReplyUser:      users[idx],
							ReplyTimeStamp: r.TimeStamp,
							ParentPost:     post,
						})

						if i < len(post.Replies)-1 {
							d = d + "\n"
						}

						threadData = threadData + d
					}
				} else if post.ParentUserId != "" {
					// Since the json files are read in chronological order,
					// it is always assumed that the relevant thread data already exists.
					parentIdx := slices.IndexFunc(threadPosts, func(threadPost ThreadPost) bool { return threadPost.ReplyTimeStamp == post.TimeStamp })

					if parentIdx == -1 || threadPosts[parentIdx].ParentPost.User.Id != post.ParentUserId {
						fmt.Println("ERROR: Found invalid thread data")
						os.Exit(1)
					}

					threadData = "Thread parent index: " + strconv.Itoa(threadPosts[parentIdx].ParentIndex)

					// remove threadPosts
					threadPosts = append(threadPosts[:parentIdx], threadPosts[parentIdx+1:]...)
				}

				for _, r := range post.Reactions {
					// add Reaction.Users
					for _, id := range r.UserIdList {
						idx := slices.IndexFunc(users, func(user User) bool { return user.Id == id })
						r.Users = append(r.Users, users[idx])
					}

					reactionUsers := getUserNameList(r.Users, func(u User) string { return u.Name })

					emojiText := ":" + r.Name + ":"

					if reactionsText != "" {
						reactionsText = reactionsText + "\n" + emoji.Sprint(emojiText+"("+strconv.Itoa(r.Count)+") - ["+strings.Join(reactionUsers, ",")+"]")
					} else {
						reactionsText = emoji.Sprint(emojiText + "(" + strconv.Itoa(r.Count) + ") - [" + strings.Join(reactionUsers, ",") + "]")
					}
				}

				row = getRow(index)
				username := post.User.Name + " (" + post.User.Id + ")"
				f.SetSheetRow(channel.Name, row, &[]interface{}{index, username, emoji.Sprint(post.Text), threadData, reactionsText, dt})

				index++
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

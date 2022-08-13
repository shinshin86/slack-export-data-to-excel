# slack-export-data-to-excel
This CLI tool of Slack export data to excel.

## Install

```sh
go install github.com/shinshin86/slack-export-data-to-excel@latest
```

## Usage

1. Download Slack export data
2. Unzip the export data.
3. Execute the command

```
slack-export-data-to-excel <export data path>
```

## Output data for Excel

* index
  * An index automatically given by the program that is assigned to each channel.
* user
  * The output is in the format `real_name(id)`.
* text
  * Posted text.
* thread
  * The information is output only if the thread exists.
    * If the post is the root of a thread, information about child threads is output.
    * If the post was made to a thread, the index of the root post is output.
* reactions
  * Reactions are output in the format `emoji(count) - [reaction users]`.
* datetime
  * The date and time of the post will be output.

## License

[MIT](https://github.com/shinshin86/slack-export-data-to-excel/blob/main/LICENSE)

## Author

[Yuki Shindo](https://shinshin86.com/en)
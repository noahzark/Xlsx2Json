# Changelog

## 1.9

### Bug fixes
- Fixed numeric dates on the 10th, 20th or 30th of a month being written as `1990-01-01`
- Cells formatted as a date in Excel are now read correctly in Date columns
- An invalid date now stops the conversion with an error naming the row and column, instead of silently writing `1990-01-01`
- Fixed a crash when using `-config=` with a config file that sets `show`; `show` is now optional

### Improvements
- Added the Date column type (yyyyMMdd, written as yyyy-MM-dd)
- Upgraded Apache POI to 5.5.1 and org.json to 20260814
- Requires Java 8 or newer
- The runnable jar is now named `xlsx2json-<version>-all.jar`
- Builds with Gradle 8 through the Gradle wrapper, with CI on GitHub Actions

## 1.8

- Fixed a bug that didn't allow a single value in an array column
- Added Time value support
- Upgraded project files
- Added support for a config file
- Fixed a null key cell bug
- Refactored the Excel parser to make the code more readable
- Only use JSONAssert at test compile time to reduce size

## 1.4

- Fixed several bugs and a minor feature update

## 1.2

- Initial release

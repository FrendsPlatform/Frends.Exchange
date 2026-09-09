# Changelog

## [2.0.0] - 2026-09-04
### Changed
- Upgraded the Task to target .NET 8.
- Removed MimeKit library.
- **Breaking:** Renamed `Options.ThrowExceptionOnFailure` to `Options.ThrowErrorOnFailure`.
- **Breaking:** Replaced `Result.ErrorMessages` with a single `Result.Error` object containing the error `Message` and `AdditionalInfo`.
### Added
- Added `Options.ErrorMessageOnFailure` to allow overriding the error message shown when the Task fails.

## [1.5.0] - 2026-05-07
### Added
- Added `Mailbox` parameter to `Connection` to allow reading emails from a mailbox different than the authenticated user. 
  For `UsernamePassword` authentication, if `Mailbox` is empty, the authenticated user's mailbox is used. For `ClientCredentials` authentication, if `Mailbox` is empty, `Input.From` is used as fallback.

## [1.4.0] - 2026-03-20
### Added
- Added DeleteReadEmails option to permanently delete emails after they are read and processed.

## [1.3.0] - 2025-10-17
### Changed
- Changed ClientSecret to a password property.

## [1.2.0] - 2024-08-22
### Changed
- Updated the Azure.Identity, Newtonsoft.Json and Mimekit libraries to their latest versions.

## [1.1.0] - 2023-05-24
### Fixed
- Skip parameter is null instead of zero when it is not supported 

## [1.0.0] - 2023-05-12
### Added
- Initial implementation
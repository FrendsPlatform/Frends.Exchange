# Changelog

## [2.0.0] - 2026-09-04
### Changed
- Target framework updated to .NET 8.
- `Exchange` task class is now static.
- Reordered `SendEmail` parameters to `Input`, `Connection`, `Options`, `CancellationToken` to follow the standard Task parameter order.
- Renamed `Options.ThrowExceptionOnFailure` to `Options.ThrowErrorOnFailure`.
### Added
- Added `Options.ErrorMessageOnFailure` to allow overriding the error message shown on failure.
- Added `Result.Error` property containing structured error details when the Task fails.

## [1.1.0] - 2024-08-22
### Changed
- Updated the MimeKit and Azure.Identity libraries to their latest versions.

## [1.0.0] - 2023-11-29
### Added
- Initial implementation
# Changelog

## [2.0.0] - 2026-09-04
### Changed
- Target framework updated to .NET 8.
- Removed MimeKit library.
- [Breaking Change] Renamed `Options.ThrowExceptionOnFailure` to `Options.ThrowErrorOnFailure`.
- [Breaking Change] Updated failure handling behavior, including exception propagation, cancellation handling, and failure message formatting.
### Added
- Added `Options.ErrorMessageOnFailure` to allow overriding the error message shown on failure.
- Added `Result.Error` property containing structured error details when the Task fails.

## [1.1.0] - 2024-08-22
### Changed
- Updated the MimeKit and Azure.Identity libraries to their latest versions.

## [1.0.0] - 2023-11-29
### Added
- Initial implementation

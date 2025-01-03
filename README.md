# VBScript IsEmpty Function Incorrectly Handles Null Values

This repository demonstrates a subtle bug in VBScript's `IsEmpty` function.  The `IsEmpty` function doesn't correctly handle `Null` values; it throws a runtime error instead of returning `True`.

## Bug Description
The `IsEmpty` function is intended to check if a variable is uninitialized or contains an empty variant.  However, if a variable is `Null`, `IsEmpty` will throw a type mismatch error.

## Reproduction
The `bug.vbs` file contains a simple function demonstrating this issue.  Running this script will result in a runtime error if the parameter 'a' is passed as `Null`.

## Solution
The `bugSolution.vbs` file provides a corrected approach using a custom function to check for both empty and `Null` values.

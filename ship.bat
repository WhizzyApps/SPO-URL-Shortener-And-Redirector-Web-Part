@if not exist node_modules\ (
  call npm install gulp-cli -g
  call npm install
)
call gulp clean
call gulp build
call gulp bundle --ship
call gulp package-solution --ship
echo.
echo.
echo Please look for the .sppkg file in subfolder .\sharepoint\solution.


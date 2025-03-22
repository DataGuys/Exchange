# Specify the file path of your JPG file
$imagePath = "Path\To\Your\Image.jpg"

# Method to convert the JPG file to a Base64 string
function Convert-ImageToBase64HTML {
    param (
        [string]$imagePath
    )

    try {
        # Read bytes of the jpg file
        $imageBytes = [IO.File]::ReadAllBytes($imagePath)
        # Convert the bytes to a Base64 encoded string
        $base64String = [Convert]::ToBase64String($imageBytes)
        # Print the formatted output for the HTML document
        $htmlBase64 = "<img src=`"data:image/jpeg;base64,$base64String`">"
        return $htmlBase64
    }
    catch {
        # Catch and show any failures during the script execution
        Write-Error "An error occurred: $($_.Exception.Message)"
    }
}

# You should replace the 'Path\To\Your\Image.jpg' with the real JPG file's location
$result = Convert-ImageToBase64HTML -imagePath $imagePath

# Outputting the final base64 HTML-ready <img> code
$result

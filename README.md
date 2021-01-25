# Slide Generator
## Description:
Create a solution that accepts user input and generates a power point slide with;<br>
Title area<br>
Text area<br>
and an image suggestion area that utilizes words in the title, and **bold** words in the text area to bring suggested images in, with ability to select multiple images to include in the slide<br>
Have them make windows form to provide the solution.

## Setup/Installation
- [Visual Studio](https://visualstudio.microsoft.com/) is needed to run this application

### Download Repo
* Clone this GitHub repository by running `git clone https://github.com/sarakane/SlideGenerator.Solution.git` in the terminal.
  * Or download the ZIP file by clicking on `Code` then `Download ZIP` from this repository.

### API KEY
This project requires a [Bing Image Search API key](https://docs.microsoft.com/en-us/bing/search-apis/bing-web-search/create-bing-search-service-resource)

To add the API key to the project create a new class in the SlideGenerator folder named `Credentials.cs`.<br>
In this class enter the following code:

```
namespace SlideGenerator
{
    internal class Credentials
    {
        public const string ApiKey = "YOUR_API_KEY_HERE";
    }
}
```
## Resources/Research:
* [Quickstart: Search for images using C# and Bing Image Search API](https://docs.microsoft.com/en-us/bing/search-apis/bing-image-search/quickstarts/rest/csharp)
* The failed solution that was provided when I was given this challenge



<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->




<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/othneildrew/Best-README-Template">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Stock Ticker Summary Table</h3>

  <p align="center">
    An explaination for the creation of a summary table from ticker infor
    <br />
    <a href="https://github.com/HsuChe/VBA_challenge"><strong>Project Github URL »</strong></a>
    <br />
    <br />
  </p>
</p>



<!-- ABOUT THE PROJECT -->
## About The Project

[![Product Name Screen Shot][product-screenshot]](https://example.com)

There is nothing more important to successful investment than to explore key business indicators to gauge growth for stocks. In this homework we are going to create a summary table for stock products specifically from 2014, 2015, and 2016

Features of the dataset:
* The dataset is divided primarily between three sheets for each of the years that are being analyzed, starting with 2014 and ending on 2016
* The ticker name, date, opening price, highest price for the day, lowest price for the day, closing price, and trade volume for the day
* The dataset is orginized in alphabetical order where the same ticker name are listed one after another. 

There homework is interested generating a few specific items for the summary table.

* The unique ticker names from the dataset
* The price change year on year
* The percentage price change year on year compared to opening price
* The total traded volume for the ticker for a given year based on the sheets. 

<!-- GETTING STARTED -->
## Obtaining list of unique ticker names and total volume traded.

To generate total volume trade, we have to create a volume counter that iterates through each row. 

* For loop on column A
  ```sh
  For i in 2 to LastRowA:
    TotalVolume = TotalVolume + Range("G" & i)
  Next i
  ```

To generate unique names we iterated through each of the rows comparing the value of the current cell in column A with the next iteration, or index + 1

* For loop on column A
  ```sh
  For i in 2 to LastRowA:
    TotalVolume = TotalVolume + Range("G" & i)
    If Range("A" & i) <> Range("A" & i + 1) Then
        TickerName = Range("A" & i)
  Next i
  ```

### Prerequisites

Some important variables we need to defined for the For loop are LastRowA, which will determine the last used cell in the column and set the range for our loop. The method I used is the following. 

<a href="https://github.com/HsuChe/VBA_challenge"><strong>Code Credit »</strong></a>

* LastRowA
  ```sh
  LastRowA = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row
  ```

We also need to generate the needed headers for the columns that our calculations will return to.

* LastRowA
  ```sh
  Range("I1") = "<ticker>"
  Range("J1") = "<yearly change>"
  Range("K1") = "<percent change>"
  Range("L1") = "<total volume>"
  ```

## Obtaining features needed to calculate price change year on year.

  * Yearly Summary Table
<br>
* 2014 Summary Table <br>
![2014 summary](Images/2014_summary.jpg)
<br>
* 2015 Summary Table <br>
![2014 summary](Images/2014_summary.jpg)
<br>
* 2016 Summary Table <br>
![2014 summary](Images/2014_summary.jpg)

To calculate price change year on year, we would need the opening price for the first iteration of an unique ticker name and the closing price for the last iteration of the same ticker ID.

The dataset is structure where the rows are ordered according to the date. We decided to compare the current iteration of ticker (i) to the next iteration (i + 1) and therefore the index (i) is going to be the last iteration of the unique value. We will be able to get the closing price for that iteration.

* Obtaining Closing Price from our index
  ```sh
  ClosingPrice = Range("F" & i)
  ```

We would also have to extract the opening price for the ticker name. We can extract the opening price for the next unique ticker name by extracting the opening price of i + 1. Because we are extracting opening price for the next iteration of unique ticker name, we would have to declare the first instance of opening price outside of the for loop.

* Obtaining Closing Price from our index
  ```sh
  Dim OpeningPrice as Double
  OpeningPrice = Range("C1")
  ```

Due to the fact that our opening price will be extracted after the current iteration of unique value is calculated, we would have to go ahead and calculate the price change year on year as well as the percentage change against the opening price.

* Store Yearly Change and Yearly Percent Change to memory
  ```sh
  YearlyChange = ClosingPrice - OpeningPrice
  ```

After the calculation is done, we can go ahead and update the opening price for the next unique value.

* Extract Opening Price
  ```sh
  OpeningPrice = Range("C" & i + 1)
  ```

## Error checking for Opening Price.
If yearly percent change uses an opening price that is 0, will draw an error for our calculation. We have to set exceptions for when opening price is 0 on hte first iteration of a unique ticker.

To do this, we would need to insert two different error checkers: 

The first error check is within the conditional when a new unique value is found and before the new opening price is updated. This error will only occur if an entire ticker name does not contain an opening price value from Jan 1 to Dec 31. 

* If OpeningPrice is 0 through every iteration for a ticker, then YearlyPercent is 0
  ``` sd
  If OpeningPrice = 0 Then
    YearlyPercent = 0
  Else
    YearlyPercent = YearlyChange / OpeningPrice
  ```

The Second will be outside the for loop as error checking needs to happen on a per row basis and update opening price to the next none 0 value when it can.

* Second error checking outside the conditional that triggers when unique value is found
  ```sh
  If OpeningPrice = 0 Then
    OpeningPrice = Range("C" & i)
  ```

## Returning all values to the correct column

We can now return the values we calculated from the if statement based on an index specifically desginated for our summary table.

* Return Values to their rightful index in the summary table 
  ```sh
  Range("I" & UniqueCounter) = TickerName
  Range("J" & UniqueCounter) = YearlyChange
  Range("K" & UniqueCounter) = YearlyPercent
  Range("L" & UniqueCounter) = TotalVolume
  ```

## Here are the results we obtained for each year starting with 2014 and ending in 2016









1. Get a free API Key at [https://example.com](https://example.com)
2. Clone the repo
   ```sh
   git clone https://github.com/your_username_/Project-Name.git
   ```
3. Install NPM packages
   ```sh
   npm install
   ```
4. Enter your API in `config.js`
   ```JS
   const API_KEY = 'ENTER YOUR API';
   ```



<!-- USAGE EXAMPLES -->
## Usage

Use this space to show useful examples of how a project can be used. Additional screenshots, code examples and demos work well in this space. You may also link to more resources.

_For more examples, please refer to the [Documentation](https://example.com)_



<!-- ROADMAP -->
## Roadmap

See the [open issues](https://github.com/othneildrew/Best-README-Template/issues) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.



<!-- CONTACT -->
## Contact

Your Name - [@your_twitter](https://twitter.com/your_username) - email@example.com

Project Link: [https://github.com/your_username/repo_name](https://github.com/your_username/repo_name)



<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements
* [GitHub Emoji Cheat Sheet](https://www.webpagefx.com/tools/emoji-cheat-sheet)
* [Img Shields](https://shields.io)
* [Choose an Open Source License](https://choosealicense.com)
* [GitHub Pages](https://pages.github.com)
* [Animate.css](https://daneden.github.io/animate.css)
* [Loaders.css](https://connoratherton.com/loaders)
* [Slick Carousel](https://kenwheeler.github.io/slick)
* [Smooth Scroll](https://github.com/cferdinandi/smooth-scroll)
* [Sticky Kit](http://leafo.net/sticky-kit)
* [JVectorMap](http://jvectormap.com)
* [Font Awesome](https://fontawesome.com)





<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/othneildrew/Best-README-Template.svg?style=for-the-badge
[contributors-url]: https://github.com/othneildrew/Best-README-Template/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/othneildrew/Best-README-Template.svg?style=for-the-badge
[forks-url]: https://github.com/othneildrew/Best-README-Template/network/members
[stars-shield]: https://img.shields.io/github/stars/othneildrew/Best-README-Template.svg?style=for-the-badge
[stars-url]: https://github.com/othneildrew/Best-README-Template/stargazers
[issues-shield]: https://img.shields.io/github/issues/othneildrew/Best-README-Template.svg?style=for-the-badge
[issues-url]: https://github.com/othneildrew/Best-README-Template/issues
[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=for-the-badge
[license-url]: https://github.com/othneildrew/Best-README-Template/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/othneildrew
[product-screenshot]: images/screenshot.png

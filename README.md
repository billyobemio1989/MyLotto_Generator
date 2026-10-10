# MyLotto_Generator
This is my Lotto Generator for PCSO.
How to use it

Open the tab for the game, for example 6/55.
Near the bottom, type the six winning numbers, like 05 12 23 34 40 41. Spaces, commas or dashes all work.
Pick the draw date. It defaults to today.
Click Save draw.

You’ll see a green message confirming it saved, and the “Latest saved winning combination” line near the top updates.

What happens behind the scenes

The draw is saved in a new file, data/NEW_DRAWS.csv. Your original CSVs stay untouched.
Every future combination for that game is checked against it, so it can never be generated again.
You can also open NEW_DRAWS.csv and fix or delete a wrong entry by hand.

Mistakes it blocks

Numbers outside the game’s range, such as 60 on 6/55
The same number typed twice
Text instead of numbers
A second draw for a date that already has one saved

One thing to watch: the “latest saved” line only shows what’s in your files, so it’s only as current as your last update. PCSO draws happen on set days, so add each result after the draw to keep it up to date.

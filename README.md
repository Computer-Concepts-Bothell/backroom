This is a HIGHLY WIP powershell tool that will probably never be useful for anyone else.
<br>
If you are intrested in Repairshopr tools Check out the inventory tool i've made, That one is much more polished and not as odd balled as this tool is designed around our weird work flow at my shop
<br>
Explaining the tool a little. 
<br>
In RS we have a custom field with every location that we would put a ticket.
To make it so we spend less time trying to remember what shelf that went on or to make sure to EVEN update it. I've made this tool that lets us scan the barcode on the ticket then scan a barcode on the shelf to update the location.
<br>
We have also had some tickets get "lost" Due to being put in the back and left "In progress" so they hadn't gotten looked at until someone found or customer called about it.
<br>
This is a little harder to set up then say the inventory tool i've made since RS saves custom vars rather oddly. 
<br>
So say a location of a spot is labed "1.1" in RS in the API calls its something like "98561"
<br>
So you will have to grab every single one of these, convert it to the formate needed for the tool to know that 98561 is 1.1 so that humans can read it. and understand that it was indeed moved to the correct spot since i have the script just text to voice the location so we don't have to look at it. 
<br> 

This tool will be updated to be a lot more friendly to other shops and ill make a tool to make the variable files a lot easier to generate. (it was painful for me)


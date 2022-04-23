V0.1:	calculation of tableaus based on brute force by random generated tableaus, text output in console
	Example: Attempts 100155155 Result: 0, 62728023, 32742080, 4528332, 155424, 1296, 0, 0
V0.6:	Working Version - creates Excel File including formulas and tableaus
V0.7:   Additional Checks: opponents & partners - no duplicates.
V0.8:	Created "Batch Runner" method - to generate multiple files in a batch without manual interaction, recursive creation of games
V1.0:	Mixed Turnier 2018.
	Issue1: uneven tableaus lead to uneven games per player. Solution: add additional round where players with too less games need to play again.
	Add game processor after initial run to add additional round (generate tableau, generate additional tableau with remaining players)
	Issue2: hard to read Excel File. Create "GUI" for tournament day (but store in excel)
	Issue4: errors when entering results - make result entry more robust (eg Confirming "winner" of game and also entering the result with validity check)
	Issue5: example: 8/6/1 template. is there really no solution with "conflict free"?" Other variant: create sets of all "combinations" of games/partners/opponents
		and then create the resulting tableaus out of this.
		Try to create all "combinations" -> creating set(menVsmen) combined with set(womenVswomen) - should be 784 items (with 8 players).
		Each round should do: remove all "future conflicting combinations (menVsmen & womenVswomen) - and repeat. Do this for all different attempts.
		This could also be easily parallized (running branches in parallel)
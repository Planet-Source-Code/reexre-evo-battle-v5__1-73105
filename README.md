﻿<div align="center">

## EVO Battle V5

<img src="PIC20111129725541810.JPG">
</div>

### Description

EVOLUTION BATTLE v5 (Neural Network trained by genetic algorithm)

Two teams; a Red one and a Cyan one fight to win battles.

Orange one is the "previous battle" best of the Red team.

Green one is the "previous battle" best of the Cyan team.

(The nearest enemy shot of the "previous battle" best fighter of the Red team is yellow circled)

At the End of each battle the team that loses Evolves (The winner team evolves with a Probability of 10%)

Each fighter have 11 Inputs and 3 Outputs:

Inputs Are:

-Enemy Left (0 means Nearest enemy is in front. 0.5 Means Nearest enemy is At 90Â° Left. 1 Means Nearest Enemy is Back)

-Enemy Right (0 means Nearest enemy is in front. 0.5 Means Nearest enemy is At 90Â° right. 1 Means Nearest Enemy is Back)

-Enemy Distance

-Enemy Shot Left (As Enemy Left but Refers to "Enemy Shot")

-Enemy Shot Right (As Enemy Right but Refers to "Enemy Shot")

-Enemy Shot Distance

-Availbale Shots

-Enemy Relative Orientation

-My Velocity

-Enemy Velocity

-Enemy Relative Moving Direction

Outputs Are:

-Rocket Left

-Rocket Right

-Fire Shot (Less than 0.5 do not fire, more than 0.5 Fire)

Each fighter have a maximum number of shots that can be fired simultaniously with a given delay. (Now it's set to 1 shot at a time)

When fired shot reach the Boundary or hit an enemy it expires.

If a fighter hits an enemy with a shot then the fighter Fitness Value is decreased.

The hitten enemy is pushed by shot, his fitness increase and it can't perform normal movement for a given time (it's frozen) and his Fitness Value is increased (In this case It's drawn with a cross).(After a fighter has been hitten it becomes "Invisible" to the enemies for a time greater than the time he can't move [frozen])

The team that have the lower avarage fitness wins. (Fitness Value Start from 8000 for each fighters)

Fitness is displayed with a line under the fighters.

If it is good (low fitness), the line go to Right for the Red team, to Left for the Cyan team.

If it is bad (high fitness), the line go to Left for the Red team, to Right for the Cyan team.

Evolution is performed by genetic algorithm that modify the genes of fighters.

The Genes are the values of a Neural Network wich have the above described Inputs & outputs.

Neural Network activity of the "previous battle" best of the Red team (Yellow one) is displaied on a picture.

Be patient, the Results of the evolution is noticiable after a long time running (a lot of generations).

Program can be stopped and restarted. At the restart (both team) last Evolved fighters are loaded. (Files: Pop1.txt and Pop2.txt) One or both of these files can be deleted to start a new random team population.

Enjoy!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2011-11-28 20:03:32
**By**             |[reexre](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/reexre.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[EVO\_Battle22156611292011\.zip](https://github.com/Planet-Source-Code/reexre-evo-battle-v5__1-73105/archive/master.zip)









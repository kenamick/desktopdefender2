#levelname('Ragious Skirmish')
#levelnum(5)
#levelduration(152000)
#leveldescbkpic(mars.bmp)
#leveldescduration(60000)
#levelbrief {
 Our sensors detect enemy ships on all sectors.
The Black League are more angry than ever.The aliens are deploying
their entire fleet into one last attack.
The time of the final battle has come.
We made it so far, only beacuse we were united.
Know that should we stop this attack, our home will be saved...
--------------------------------------------------------------------
Mission: The Earth must survive...
}

#levelstate(1,10000) {
#createbs(1)
#timewarpship(4900,6000,random,1,1)
#warpship(0,1,3)
#warpship(1,1,3)
#warpship(0,2,2)
#warpship(1,2,3)
}


#levelstate(2,35000) {
#warpship(0,1,3)
#warpship(1,2,2)
#timewarpship(6000,6500,random,3,1)
#timewarpship(2300,2600,random,2,1)
#timewarpmeteor(4800,6050,2,8)
#timewarpmeteor(7500,8000,1,random)
}

#levelstate(3,15000) {
}


#levelstate(4,8000) {
#addsmq(8000,,,,,,Prepare for massive attack.)
}


#levelstate(5,50000) {
#createbs(1)
#timewarpship(7500,8000,random,2,2)
#timewarpship(7500,7950,random,3,1)
#timewarpship(7000,8000,random,1,1)
#timewarpship(24500,24600,random,4,1)
#timewarpmeteor(9000,10000,2,random)
}

// regeneration pause
#levelstate(6,10000) {
}

// FINAL
#levelstate(7,8000) {
#addsmq(5000,,,,,,We shall never surrender... )
}

#levelstate(8,6969) {
#warpship(0,1,6)
#warpship(0,2,6)
#warpship(1,1,6)
#warpship(1,2,6)
#warpship(0,3,3)
#warpship(0,4,3)
#warpship(1,3,1)
#warpship(1,4,1)
}

#levelstate(9,9690) {
#givebonus(0)
}


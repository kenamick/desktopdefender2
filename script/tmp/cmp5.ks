#levelname(Ragious Skirmish)
#levelnum(5)
#levelduration(150000)
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 The time of the total skirmish has come....
 It's now or never 
 Enemy forces scout on all sides.Massive attacks...
}

#levelstate(1,10000) {
#createbs(1)
#timewarpship(4900,6000,random,1,1)
#warpship(0,1,3)
#warpship(1,1,3)
#warpship(0,2,3)
#warpship(1,2,3)
}


#levelstate(2,35000) {
#warpship(0,1,3)
#warpship(1,2,2)
#timewarpship(5500,6000,random,3,1)
#timewarpship(2100,2200,random,2,1)
#timewarpmeteor(4800,6050,2,8)
#timewarpmeteor(7500,8000,1,random)
}

#levelstate(3,15000) {
}


#levelstate(4,8000) {
#addsmq(8000,,,,,,Enemies are preparing their last ragios attack! )
}


#levelstate(5,50000) {
#createbs(1)
#timewarpship(7500,8000,random,2,2)
#timewarpship(6700,6950,random,3,1)
#timewarpship(7500,8000,random,1,2)
#timewarpship(24500,24600,random,4,1)
#timewarpmeteor(9000,10000,2,random)
}

// regeneration pause
#levelstate(6,10000) {
}

// FINAL
#levelstate(7,6000) {
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


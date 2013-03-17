#authorinfo 
 .name Pro-XeX
 .date 16.04.2003  

#levelname(Early Attack)
#levelnum(1)
#levelduration(121000)
// setup briefing
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
  The Black League scout forces came faster than we expected.
 The enemies warped around 5:39 OET and engaged our planet.
 The war has began...
}

// LEVEL STATES

#levelstate(1,3000) {
#addsmq(3000,,,,,,Enemies are closing up!)
}

// p1-pos(1r,2l,type,num)
// initial attacks
#levelstate(2,12500) {
#warpship(random,1,9)
}

#levelstate(3,12500) {
#warpship(1,2,9)
#timewarpship(5000,5350,1,2,2)
}

#levelstate(4,4500) {
#addsmq(4500,,,,,,Warning: 2nd attack wave!)
}

#levelstate(5,20000) {
#timewarpship(4900,5100,random,1,3)
#timewarpship(7000,8000,random,2,2)
}

#levelstate(6,10000) {
#warpship(1,1,5)
}

#levelstate(7,4500) {
#addsmq(4500,,,,,,Prepare for final attack wave!)
#warpship(0,1,1)
}

#levelstate(8,38000) {
#timewarpship(3700,4000,random,1,2)
#timewarpship(4800,5200,random,2,2)
}

#levelstate(9,10000) {
#addsmq(10000,,,,,,Initial attack has been repulsed.Good job!)
}





#authorinfo 
 .name Pro-XeX
 .date 17.04.2003  

#levelname('Meteor shower')
#levelnum(2)
#levelduration(130900)
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 The enemies are preparing for unseen attack.Our scouts report that
meteors from the first asteroid ring have been redirected towards
the Earth.
Do not let any of those meteors to hit the Earth, damage will be
enormous.
The Zonerian Central Intelligence reported for large Motherships,
be on alert!
In case things get ugly, the Zonerians are ready to unleash
annihilation bomb.
--------------------------------------------------------------------
Mission: Divert the meteor attack!
--------------------------------------------------------------------
Hint: Use guided missiles when targeting meteors!
}

// LEVEL STATES

#levelstate(1,4800) {
#addsmq(4800,,,,,,Warning!Meteors are closing.)
}

#levelstate(2,25000) {
#timewarpmeteor(3050,4050,1,random)
}

#levelstate(3,20000) {
// pause
}

#levelstate(4,10000) {
#addsmq(8000,,,,,,Alert!Close meteor attack. )
#warpmeteor(2,4)
#warpmeteor(2,8)
}

#levelstate(5,8000) {
#warpmeteor(2,4)
#warpmeteor(1,4)
#warpmeteor(2,8)
}

#levelstate(6,5000) {
#addsmq(5000,,,,,,Alert!Enemies detected.)
}

#levelstate(7,9000) {
#warpship(0,2,4)
#warpship(1,1,3)
}

#levelstate(8,4000) {
#addsmq(4000,,,,,,Motherships closing!)
}

#levelstate(9,10000) {
#timewarpship(5000,5200,random,3,1)
#timewarpship(4500,4800,random,1,1)
}

#levelstate(10,22500) {
#addsmq(10000,,,,,,12 minutes till annihilation bomb!  )
#warpship(random,3,1)
#timewarpship(5200,5800,random,1,1)
#timewarpship(6900,7000,random,3,1)
#timewarpship(4500,4900,random,2,1)
}

#levelstate(11,10000) {
#givebonus(0)
}

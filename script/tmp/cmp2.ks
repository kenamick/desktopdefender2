#authorinfo 
 .name Pro-XeX
 .date 17.04.2003  

#levelname(Meteor Shower)
#levelnum(2)
#levelduration(130500)
// setup briefing
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
  The enemies are playing hard deploying meteors from the mars orbit.
 Defend the Earth at any cost!!!
 Use missiles to destroy the meteors!
}

// LEVEL STATES

#levelstate(1,4000) {
#addsmq(4000,,,,,,Warning: Meteors approaching! )
}

#levelstate(2,25000) {
#timewarpmeteor(3050,4050,1,random)
}

#levelstate(3,20000) {
// pause
}

#levelstate(4,10000) {
#addsmq(8000,,,,,,Warning: Close meteor attack!Bunker alert! )
#warpmeteor(2,4)
#warpmeteor(2,8)
}

#levelstate(5,8000) {
#warpmeteor(2,4)
#warpmeteor(1,4)
#warpmeteor(2,8)
}

#levelstate(6,5000) {
#addsmq(5000,,,,,,Alert: Incoming enemy ships! )
}

#levelstate(7,9000) {
#warpship(0,2,4)
#warpship(1,1,5)
}

#levelstate(8,4000) {
#addsmq(4000,,,,,,Carrier ships detected! )
}

#levelstate(9,10000) {
#timewarpship(5000,5200,random,3,1)
#timewarpship(3900,4000,random,1,1)
}

#levelstate(10,22500) {
#addsmq(10000,,,,,,Annihilate missile shall be deployed in 15 seconds. )
#warpship(random,3,1)
#timewarpship(4000,4500,random,1,2)
#timewarpship(6500,7500,random,3,1)
#timewarpship(3500,4000,random,2,1)
}

#levelstate(11,10000) {
#givebonus(0)
}

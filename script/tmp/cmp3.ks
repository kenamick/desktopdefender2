#levelname(Beam Disease)
#levelnum(3)
#levelduration(122000)
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 Beam Carriers are warping...
}

// LEVEL STATES

#levelstate(1,15000) {
#warpship(0,1,3)
#warpship(1,2,2)
#warpship(1,1,2)
#warpship(0,2,2)
}

#levelstate(2,10000) {
#addsmq(10000,,,,,,Warning: Large Carrier Ships detected! )
#timewarpship(4900,5200,random,2,2)
}

#levelstate(3,40000) {
#timewarpship(10300,11100,random,4,1)
#timewarpship(5500,6000,random,1,1)
}

#levelstate(4,20000) {
#timewarpmeteor(4000,5000,2,random)
#timewarpmeteor(9700,10100,1,random)
#timewarpship(4200,4900,random,1,1)
#timewarpship(5200,5700,random,2,1)
}

#levelstate(5,22000) {
}

#levelstate(6,12000) {
#addsmq(4000,,,,,,Mission completed.)
#addsmq(4000,,,,,,Nice job defender.)
}
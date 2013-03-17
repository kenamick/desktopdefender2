#authorinfo 
 .name Peter "Pro-XeX" Georgiev
 .date 12.09.2002  

#levelname(Early Attack)
#levelnum(1)
#levelduration(156000)
// setup briefing
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 We shall fight to the last piece of metal....
}

// LEVEL STATES

#levelstate(1,1500) {
// add message to the cockpit scrolling-message-queue
#addsmq(1500,,,,,,Here we go...)
#createbs()
#warpship(1,4,1)
}

// initial attacks
#levelstate(2,20000) {
#timewarpship(2800,3200,random,1,2)
}

#levelstate(3,5000) {
#addsmq(5000,,,,,,Warning...meteor attack...)
}

// the meteor shower
#levelstate(4,25000) {
#warpmeteor(2,4)
#warpmeteor(2,8)
#timewarpmeteor(4000,4400,1,random)
}

#levelstate(5,12000) {
//
#addsmq(10000,,,,,,Large Carrier ships detected, prepare for a massive alien warp...)
}

#levelstate(6,29000) {
#addsmq(7000,,,,,,Destroy the carrier and the interceptors are gone too.)

#timewarpship(1400,2200,random,2,1)
#timewarpship(5800,8800,random,3,1)
}

// prepare bonus
#levelstate(7,6000) {
#timewarpship(1100,2000,random,2,1)
#addsmq(6000,,,,,,Allies are deploying an annihilate missile!)
}

// annihilate 
#levelstate(8,8000) {
#givebonus(0)
}

// closing message
#levelstate(10,12000) {
#addsmq(10000,,,,,,Congrats!You are a...worthy defender! ;))
}

// total end
#levelstate(10,10000) {
#destroybs()
#destroyallbunkers()
}



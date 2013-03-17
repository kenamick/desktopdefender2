#authorinfo 
 .name Peter "Pro-XeX" Georgiev
 .date 12.09.2002  

// THIS IS NOT C ;)
// PLEASE RESPECT THIS SCRIPT SYNTAX OTHERWISE ERRORS SHALL FILL YOUR PCBOX

#levelname(Early Attack)
#levelnum(1)
#levelduration(156000)
// setup briefing
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 The Black League scout forces came faster than we expected.
 The enemies warped around 5:39 OET and engaged the Earth.
 The war has began...
}

// LEVEL STATES

#levelstate(1,8000) {
// add message to the cockpit scrolling-message-queue
#addsmq(6000,,,,,,Alien fleet warping to the right in 3 second...)
}

// initial attacks
#levelstate(2,15000) {
#timewarpship(2000,2500,1,2,2)
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

#levelstate(5,20000) {
//
#addsmq(10000,,,,,,Large Carrier ships detected, prepare for a massive alien warp...)
}

#levelstate(6,29000) {
#addsmq(7000,,,,,,Destroy the carrier and the interceptors are gone too.)

#timewarpship(1400,2200,random,2,1)
#timewarpship(5800,8800,random,3,1)
}

// prepare bonus
#levelstate(7,7000) {
#timewarpship(1100,2000,random,2,1)
#addsmq(7000,,,,,,Allies are deploying an annihilate missile!)
}

// annihilate 
#levelstate(7,8000) {
#givebonus(0)
}

// closing message
#levelstate(10,19000) {
#addsmq(10000,,,,,,Congrats!You are a...worthy defender! ;))
}



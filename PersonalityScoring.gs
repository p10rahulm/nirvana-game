

Logger.log(Date.now() )

function EmailFormConfirmation() {
  
  // 3-Step Instructions:
  // 1) Place the following formula into an empty cell in the first row, 
  //     in a Column to the right of Columns that populate from the form:   =indirect("C"&counta(A1:A))
  //     Change C in above formula to the Column in your sheet that has email addresses from the form 

  // 2) In the following two program lines, "var sheetname" and "var columnnumber",
  // Change the name of the sheet (currently set to Sheet1, the default name)
  // Change the column number to the number representing the Column that you placed the indirect formula, 
  // from step 1, above
  // i.e. B=2, C=3, D=4, J=10, O=15, T=20, Z=26, AA=27, AZ=52, etc...(currently set to 10, for Column J)
  
  var sheetname = "FormResponses1"
  var columnnumber = 42



//--------------------------------------------------------------------------------------------------------------------------------------
// Variables declaration
//--------------------------------------------------------------------------------------------------------------------------------------
  var gender = 4
  var maritalstatus = 7
  var fashion = 12
  var doctortype = 13
  var mobileaddict = 15
  var ideadiffusion = 39
  var bollywoodmusic = 20
  
  var petscore = 16
  var subject = 24
  var interestindiffsports = 28
  var politicalthought = 25
  var religionmeat = 31
  var demonetization = 32
  var moab = 34
  var debate = 37
  
  var education = 5
  var highestdegree = 6
  var describestyle = 40
  var smartphone = 14
  var sharedcab  = 41
  var socialscore = 38
  var internalexternal = 36
  var decisionlaziness = 35
  var truepatriotism = 33
  var financialrisk = 29
  var indoorsport = 27
  var agegroup = 3
  var vacation = 19
  var physicalsport = 26
  var numlanguages = 10
  
  var bookreadertype = 23
  var tvtype = 22
  var musiccount = 21
  var partofday = 17
  var animallover = 16
  var mobileaddict = 15
  var faithvsscience = 13
  
  var ditcher = 11
  
//--------------------------------------------------------------------------------------------------------------------------------------

  
  
  
  
  // 3) After saving this script, select "Triggers" => "Current script's triggers" => Click to add a script
  //     Choose this script's name, select "From spreadsheet" and select "On form submit", and then save 
  // 
  //  That's it!


//--------------------------------------------------------------------------------------------------------------------------------------
// Main Script
//--------------------------------------------------------------------------------------------------------------------------------------



  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(sheetname);
  var email = sheet.getRange(1,columnnumber).getValue();  
  
  // Determines row number of most recent form submission and sets it as "lastrow"
  if (sheet.getRange(sheet.getMaxRows(),1).getValue() != "") {
    var lastrow = sheet.getMaxRows()
        } 
  else {
      var count = 0
      for (var i = 0; i < sheet.getMaxRows(); i++) {
        if (sheet.getRange(sheet.getMaxRows()-i,1).getValue() != "") {
          var lastrow = sheet.getMaxRows()-i
          break;
        }  
      }
   }


//--------------------------------------------------------------------------------------------------------------------------------------
// Variables declaration
//--------------------------------------------------------------------------------------------------------------------------------------
  
  // Do some checks. Check the logs in >view>logs
  Logger.log("email is: ")
  Logger.log(email)
  Logger.log("lastrow is: ")
  Logger.log(lastrow)
  
  
  // Email address regex (regular expression)
  // Test for valid Email Pattern/Format - Allows any two-letter country code top level domain, 
  // and only specific generic top level domains 
  // (update via: http://en.wikipedia.org/wiki/List_of_top_level_domains)  

  var emailPattern = /^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+(?:[A-Z]{2}|aero|asia|biz|com|coop|edu|gov|info|int|jobs|mil|mobi|name|museum|name|net|org|pro|tel|travel)\b/;
  
  var validEmailAddress = emailPattern.test(email); 
  Logger.log(validEmailAddress)
  // The following sends an email if the email pattern is valid (i.e. if the email address is of an acceptable format)  
  // Edit the text you want in the body of the email and the subject you want (send a test message to yourself to test)
                                    

  
    if (validEmailAddress == true) {
      Logger.log("inside")
      
//--------------------------------------------------------------------------------------------------------------------------------------
// Create the required functions
//--------------------------------------------------------------------------------------------------------------------------------------
      
      
//--------------------------------------------------------------------------------------------------------------------------------------
// Create the random function
//--------------------------------------------------------------------------------------------------------------------------------------
/**
 * Returns a random integer between min (inclusive) and max (inclusive)
 * Using Math.round() will give you a non-uniform distribution!
 */
function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}




//--------------------------------------------------------------------------------------------------------------------------------------
// Load images onto an array
//--------------------------------------------------------------------------------------------------------------------------------------
 var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
 Logger.log("Remaining email quota: " + emailQuotaRemaining);
 
// array of images
var images = [];

// push ten images to the array
//tree
images.push("https://cdn.pixabay.com/photo/2014/05/05/14/14/tree-338211_960_720.jpg");
//man on cycle
images.push("https://cdn.pixabay.com/photo/2014/10/03/09/17/man-471192_960_720.jpg");
//man on cycle
images.push("https://cdn.pixabay.com/photo/2016/11/26/12/00/kyoto-1860521_960_720.jpg");
//sun and sky behind clouds
images.push("https://cdn.pixabay.com/photo/2013/10/02/23/03/dawn-190053_960_720.jpg");
//yosemite behind snow
images.push("https://www.nature.org/cs/groups/webcontent/@web/documents/media/2016-photocontest-yosemite-w-1.jpg");
//sun and sky and everything
images.push("https://cdn.pixabay.com/photo/2017/02/22/20/02/landscape-2090495_960_720.jpg");
//path through the woods
images.push("https://cdn.pixabay.com/photo/2015/03/26/09/56/trail-690619_960_720.jpg");
//autumn paths
images.push("http://maxpixel.freegreatpicture.com/static/photo/1x/Forest-Autumn-Fall-Season-Nature-Road-Landscape-1072823.jpg");
//dense forest
images.push("https://cdn.pixabay.com/photo/2016/04/22/13/16/forest-1345747_960_720.jpg");
//colourful
images.push("http://maxpixel.freegreatpicture.com/static/photo/1x/Woods-Tree-Leaves-Fall-Nature-Autumn-Red-Season-1072821.jpg");
//colourful
images.push("http://maxpixel.freegreatpicture.com/static/photo/1x/Woods-Tree-Leaves-Fall-Nature-Autumn-Red-Season-1072821.jpg");

//baby for those who left
images.push("http://vignette2.wikia.nocookie.net/callofduty/images/5/53/We_Want_You.jpg/revision/latest?cb=20111212162319");


// output
Logger.log(sheet.getRange(lastrow,ditcher).getValue());
//Logger.log(getRandomInt(1,10));
//Logger.log(images);
//Logger.log(images[11]);
//--------------------------------------------------------------------------------------------------------------------------------------
// For those who didn't fill the full form
//--------------------------------------------------------------------------------------------------------------------------------------


if (sheet.getRange(lastrow,ditcher).getValue() == "No") {
Logger.log("I'm inside this place");
var imageUrl = images[11];

// Create the blob to be embedded in email   
var imageBlob = UrlFetchApp
                          .fetch(imageUrl)
                          .getBlob()
                          .setName("personalityImageBlob");

//create a generic subject
var mailSubject = "Hope to have you back soon!";
//create the boilerplate html
var message = "<br><img src='cid:personalityImage'; style='width:100%'><br>";
message = message.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;">');
message = message.concat('To take <a href = "https://goo.gl/forms/ImhuWGgHJ743c5c82">this</a> quiz again. Be back when you can!<br></p>')
      
                          
   MailApp.sendEmail({
     to: "rahul.madhavan@atidiv.com",
     //to: email,
     bcc:"rahul.madhavan@atidiv.com,ankit.baraskar@atidiv.com",
     subject: mailSubject,
     htmlBody: message,
     inlineImages:
       {
         personalityImage: imageBlob,
       }
   });
   
}
else {

//--------------------------------------------------------------------------------------------------------------------------------------
// For those who filled the full form
//--------------------------------------------------------------------------------------------------------------------------------------


Logger.log("I'm inside the other place");
// Create an image url
var imageUrl = images[getRandomInt(0,10)]

// Create the blob to be embedded in email   
var imageBlob = UrlFetchApp
                          .fetch(imageUrl)
                          .getBlob()
                          .setName("personalityImageBlob");

// Get Social Share images
var facebookBlob = UrlFetchApp
                          .fetch('https://cache.addthiscdn.com/icons/v3/thumbs/32x32/facebook.png')
                          .getBlob()
                          .setName("facebookImageBlob");
var gplusBlob = UrlFetchApp
                          .fetch('https://cache.addthiscdn.com/icons/v3/thumbs/32x32/google_plusone_share.png')
                          .getBlob()
                          .setName("gplusImageBlob");

var twitterBlob = UrlFetchApp
                          .fetch('https://cache.addthiscdn.com/icons/v3/thumbs/32x32/twitter.png')
                          .getBlob()
                          .setName("twitterImageBlob");

var whatsappBlob = UrlFetchApp
                          .fetch('https://cache.addthiscdn.com/icons/v3/thumbs/32x32/whatsapp.png')
                          .getBlob()
                          .setName("whatsappImageBlob");

var emailBlob = UrlFetchApp
                          .fetch('https://cache.addthiscdn.com/icons/v3/thumbs/32x32/email.png')
                          .getBlob()
                          .setName("emailImageBlob");





//--------------------------------------------------------------------------------------------------------------------------------------
// Create the message
//--------------------------------------------------------------------------------------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------

//create a generic subject
var mailSubject = "The story of you*";
//create the boilerplate html
var message = "<br><img src='cid:personalityImage'><br><hr>"

//Create the various categories of messages

 Logger.log("hihihi".search("hi")>=0);
 Logger.log("hihihi".search("bo")>=0);

//--------------------------------------------------------------------------------------------------------------------------------------
// Personality category
//--------------------------------------------------------------------------------------------------------------------------------------

var personalityscore = 0

if (sheet.getRange(lastrow,maritalstatus).getValue() == "Married"){
  personalityscore +=10
}
if (sheet.getRange(lastrow,gender).getValue() == "Female"){
  personalityscore +=5
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Dating/Matchmaking apps")>=0 ){
  personalityscore -=20
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Dating/Matchmaking apps")>=0 && sheet.getRange(lastrow,maritalstatus).getValue() == "Married"){
  personalityscore -=50
}
if (sheet.getRange(lastrow,subject).getValue() == "English/Other Language/Literature Studies" ){
  personalityscore -=10
}
if (sheet.getRange(lastrow,politicalthought).getValue() == "Samrat Ashoka" ){
  personalityscore +=20
}
if (sheet.getRange(lastrow,interestindiffsports).getValue().search("WWE / UFC")>=0 ){
  personalityscore -=20
}
if (sheet.getRange(lastrow,truepatriotism).getValue() == "Yes, doing this will make Indians more disciplined and respect their country"){
  personalityscore +=20
}
if (sheet.getRange(lastrow,demonetization).getValue() == "It was a great move with great results, but the common man had to suffer"){
  personalityscore +=15
}
if (sheet.getRange(lastrow,moab).getValue() == "That was a wrong move... innocents could have gotten killed" ){
  personalityscore +=15
} else if (sheet.getRange(lastrow,moab).getValue() == "It was a necessary move even if there was a 1% chance of success" ){
  personalityscore -= 30
}

if (sheet.getRange(lastrow,petscore).getValue() == "I don't care about domestic animals"){
  personalityscore -=5
} else if (sheet.getRange(lastrow,petscore).getValue() == "Animals = food"){
  personalityscore -=5
} else if (sheet.getRange(lastrow,petscore).getValue() == "I don't have a pet, but would like one"){
  personalityscore +=10
} else if (sheet.getRange(lastrow,petscore).getValue() == "I have a pet"){
  personalityscore +=20
} else if (sheet.getRange(lastrow,petscore).getValue() == "My house is a petting zoo"){
  personalityscore +=20
} 

if (sheet.getRange(lastrow,religionmeat).getValue() == "It is disgusting and it offends me if anyone eats it"){
  personalityscore +=0
} else if (sheet.getRange(lastrow,religionmeat).getValue() == "I am a vegetarian but not for religious reasons"){
  personalityscore +=35
} else if (sheet.getRange(lastrow,religionmeat).getValue() == "I personally don't eat it, but I don't care if anyone else does or not"){
  personalityscore +=25
} else if (sheet.getRange(lastrow,religionmeat).getValue() == "I eat either, so it doesn't matter to me"){
  personalityscore +=15
} 


if (sheet.getRange(lastrow,debate).getValue() == "To refute all of your opponent's points"){
  personalityscore -=30
} else if (sheet.getRange(lastrow,debate).getValue() == "To win at any cost"){
  personalityscore -=10
} else if (sheet.getRange(lastrow,debate).getValue() == "To reach an agreement"){
  personalityscore +=30
} 

if (sheet.getRange(lastrow,sharedcab).getValue() == "Wonder why you got into this cab and spoiled your mood"){
  personalityscore +=0
} else if (sheet.getRange(lastrow,sharedcab).getValue() == "Nothing - I don't want to interfere in a stranger's problem."){
  personalityscore +=10
} else if (sheet.getRange(lastrow,sharedcab).getValue() == "Offer them advice on their problem you overheard"){
  personalityscore +=20
} else if (sheet.getRange(lastrow,sharedcab).getValue() == "Talk to the driver and try to change the mood"){
  personalityscore +=30
} else if (sheet.getRange(lastrow,sharedcab).getValue() == "Explicitly ask them if they're okay"){
  personalityscore +=40
} 

sharedcab

Logger.log("personalityscore");
Logger.log(personalityscore);


personalitymsg = '';
personalitymsg = personalitymsg.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
personalitymsg = personalitymsg.concat('Your Personality</p>');

if(personalityscore >=110) { 
personalitymsg = personalitymsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
personalitymsg = personalitymsg.concat('Zeus - The Voice Of Reason<br></p>');
personalitymsg = personalitymsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
personalitymsg = personalitymsg.concat("You are the voice of conscience in your group of friends. You have a very clear definition of right and wrong, and a high level of empathy for humans and animals alike. When it comes to making the right decisions, you\'re tough to beat, but sometimes the right decision might not be the best decision for you personally - too much of self-sacrifice can be injurious to your health");

} else if (personalityscore >=80) { 

personalitymsg = personalitymsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
personalitymsg = personalitymsg.concat('Odin - The Wise Judge<br></p>');
personalitymsg = personalitymsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
personalitymsg = personalitymsg.concat("You are the final authority when it comes to disputes between your friends or relatives. Not only that - you also silently judge people without ever talking to them based on what they wear, how they speak and a variety of other things. Occasionally, you might just zone out and pass an arbitrary decision or two, but for the most part your sharp observations can shut a difficult case pretty quick.");

} else if (personalityscore >=0) { 

personalitymsg = personalitymsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
personalitymsg = personalitymsg.concat('Loki - The Devil\'s Advocate<br></p>');
personalitymsg = personalitymsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
personalitymsg = personalitymsg.concat("You never make judgments, but try to reach agreements instead - you argue both sides, sit on the fence, and occasionally even just sit silent and watch the fun of a fight unfolding in front of your very eyes. Oh - and you make sure you derive some profits from all of this. This is not to say that you\'re a terrible person - you just have a hazier definition of morality than most people, and it works just fine for you.");

} else { 

personalitymsg = personalitymsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
personalitymsg = personalitymsg.concat('Thor - The Outrageous Outlaw<br></p>');
personalitymsg = personalitymsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
personalitymsg = personalitymsg.concat("A born rebel, you don\'t think much of society\'s definitions of right and wrong. You make your own rules based on your own experiences. You aren\'t afraid to speak out against authority figures, which might prove to be troublesome in the long run. But for you, it\'s a small price to pay for freedom.");

}





personalitymsg = personalitymsg.concat('<br><hr></p>');
message = message.concat(personalitymsg);

//--------------------------------------------------------------------------------------------------------------------------------------
// Love category - done
//--------------------------------------------------------------------------------------------------------------------------------------
 

//--------------------------------------------------------------------------------------------------------------------------------------
//Create a love score
var lovescore = 0

if (sheet.getRange(lastrow,maritalstatus).getValue() == "Single"){
  lovescore +=10
}
if (sheet.getRange(lastrow,gender).getValue() == "Male"){
  lovescore +=10
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Dating/Matchmaking apps")>=0 ){
  lovescore +=20
}
if (sheet.getRange(lastrow,bollywoodmusic).getValue() == "Jagjit Singh" ){
  lovescore +=20
}
if (sheet.getRange(lastrow,physicalsport).getValue().search("Yoga/zumba/dance based exercises")>=0){
  lovescore +=40
}


if (sheet.getRange(lastrow,socialscore).getValue() == "10-100"){
  lovescore +=0
} else if (sheet.getRange(lastrow,socialscore).getValue() == "100-500"){
  lovescore +=5
} else if (sheet.getRange(lastrow,socialscore).getValue() == "500-1000"){
  lovescore +=15
} else if (sheet.getRange(lastrow,socialscore).getValue() == "> 1000"){
  lovescore +=25
} 
if (sheet.getRange(lastrow,bookreadertype).getValue().search("Rumi")>=0 ){
  lovescore += 25
} 



if (sheet.getRange(lastrow,fashion).getValue() == "I can never find the right outfit"){
  lovescore +=0
} else if (sheet.getRange(lastrow,fashion).getValue() == "I keep it basic - whatever everyone else seems to wear"){
  lovescore +=10
} else if (sheet.getRange(lastrow,fashion).getValue() == "I follow the latest trends in fashion"){
  lovescore +=20
} else if (sheet.getRange(lastrow,fashion).getValue() == "I can pull off anything that comes to my hand"){
  lovescore +=25
} else if (sheet.getRange(lastrow,fashion).getValue() == "I create my own trends"){
  lovescore +=40
} 
if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'm pretty sure I won't buy it"){
  lovescore +=0
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "My friends are likely to buy it before me"){
  lovescore +=5
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll believe it when i see it"){
  lovescore +=10
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll wait for the reviews"){
  lovescore +=15
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I can't wait to get my hands on it"){
  lovescore +=25
} 
Logger.log("LoveScore");
Logger.log(lovescore);


//--------------------------------------------------------------------------------------------------------------------------------------
lovemessage = '';
lovemessage = lovemessage.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
lovemessage = lovemessage.concat('Your Love Story</p>');

if(sheet.getRange(lastrow,mobileaddict).getValue().search("Dating/Matchmaking apps")>=0 && sheet.getRange(lastrow,maritalstatus).getValue() == "Married"){

lovemessage = lovemessage.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lovemessage = lovemessage.concat('The Commitment Phobic<br></p>');
lovemessage = lovemessage.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lovemessage = lovemessage.concat("You\’ve been in love more than once - and the last time should have been the last. But in spite of tying the knot or making a commitment, you\’re confused by the presence of a stronger attraction, and it\'s keeping you up at night. It\’s a strange time for you, but hopefully you\’ll find the right answers before things turn even stranger.");

} else if (lovescore <40) {

lovemessage = lovemessage.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lovemessage = lovemessage.concat('The Workaholic Hermit<br></p>');
lovemessage = lovemessage.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lovemessage = lovemessage.concat("Your work is your devotion and the office is your temple. You realize the value of time and want to make the most of the little that there is. You don\'t let external factors affect your capability - including headaches from relationships. This self-control, devotion and the zeal to progress professionally helps you in being the best at what you do. You\'re a great friend but there\'s no time for love. Love is your kryptonite, and you\'ve left it behind.. far far away..");

} else if (lovescore <70) {

lovemessage = lovemessage.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lovemessage = lovemessage.concat('The Pragmatic Persona<br></p>');
lovemessage = lovemessage.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lovemessage = lovemessage.concat("It might be sunny outside, but your imagination is clouded over with a sense of longing. Either you\’ve had tough luck so far, or you have a generally pessimistic view when it comes to romance. This causes you to be guarded and closed off when it comes to matters of the heart, but you might want to relax and open it up for the right person. ");

} else if (lovescore <130) {

lovemessage = lovemessage.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lovemessage = lovemessage.concat('The Hapless Romantic<br></p>');
lovemessage = lovemessage.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lovemessage = lovemessage.concat("Whether you\’ve faced one rejection or many (or none), your attitude barely differs when you try the next time around. You\’ve either consumed too much of 90s Bollywood, and imagine you\’ll find your partner when you\’re both carrying a stack of books and bump into each other, or you\’ve seen too much of Friends and expect relationships to keep getting complicated over time.");

} else {

lovemessage = lovemessage.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lovemessage = lovemessage.concat('The Compulsive Casanova<br></p>');
lovemessage = lovemessage.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lovemessage = lovemessage.concat(" Your personality is as dashing as can be. You get noticed among the crowd and the attention attracts the best candidates towards you. Your style evolves ahead of its time; glamour and fame comes to you naturally. But you can\'t survive on curd rice everyday and like to have some variety. The same applies to your love life. Why settle for curd rice when you can have Italian Lasagna...");


}




lovemessage = lovemessage.concat('<br><hr></p>');
message = message.concat(lovemessage);



//--------------------------------------------------------------------------------------------------------------------------------------
// Career category - done
//--------------------------------------------------------------------------------------------------------------------------------------
var careerscore = 0

if (sheet.getRange(lastrow,education).getValue() == "Science" || sheet.getRange(lastrow,education).getValue() == "Engineering" || sheet.getRange(lastrow,education).getValue() == "Law" || sheet.getRange(lastrow,education).getValue() == "Management"){
  careerscore +=10
}
if (sheet.getRange(lastrow,gender).getValue() == "Male"){
  careerscore +=5
}
if (sheet.getRange(lastrow,highestdegree).getValue() == "Master's"){
  careerscore +=10
}
if (sheet.getRange(lastrow,highestdegree).getValue() == "Doctorate"){
  careerscore +=10
}
if (sheet.getRange(lastrow,describestyle).getValue() == "A watch" || sheet.getRange(lastrow,describestyle).getValue() == "My phone"){
  careerscore +=20
}
if (sheet.getRange(lastrow,smartphone).getValue() == "Apple" || sheet.getRange(lastrow,smartphone).getValue() == "Google Pixel"){
  careerscore +=20
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Productivity apps")>=0){
  careerscore +=20
}
if (sheet.getRange(lastrow,tvtype).getValue()=="News"){
  careerscore +=20
}
if (sheet.getRange(lastrow,financialrisk).getValue()=="I invest in my present"){
  careerscore -=20
}

if (sheet.getRange(lastrow,decisionlaziness).getValue()=="What aadhar card?"){
  careerscore +=0
} else if (sheet.getRange(lastrow,decisionlaziness).getValue()=="2017"){
  careerscore +=5
} else if (sheet.getRange(lastrow,decisionlaziness).getValue()=="2013-16"){
  careerscore +=10
} else if (sheet.getRange(lastrow,decisionlaziness).getValue()=="2012"){
  careerscore +=10
} else if (sheet.getRange(lastrow,decisionlaziness).getValue()=="2011"){
  careerscore +=17
} else if (sheet.getRange(lastrow,decisionlaziness).getValue()=="2010"){
  careerscore +=20
} 

if (sheet.getRange(lastrow,internalexternal).getValue()=="I don't really know, things just happened" || sheet.getRange(lastrow,internalexternal).getValue()=="My parents/family/better half asked me to"){
  careerscore +=0
} else if (sheet.getRange(lastrow,internalexternal).getValue()=="I like the money."){
  careerscore +=5
} else if (sheet.getRange(lastrow,internalexternal).getValue()=="What I do at work allows me to do what I want outside work." || sheet.getRange(lastrow,internalexternal).getValue()=="I don't know where I want to go, and this profession buys me time to think about it"){
  careerscore +=10
} else if (sheet.getRange(lastrow,internalexternal).getValue()=="I know where I want to go and this profession allows me to get there"){
  careerscore +=25
} 

if (sheet.getRange(lastrow,debate).getValue()=="To reach an agreement"){
  careerscore +=0
} else if (sheet.getRange(lastrow,debate).getValue()=="To be logically consistent with what you started"){
  careerscore +=5
} else if (sheet.getRange(lastrow,debate).getValue()=="To refute all of your opponent's points"){
  careerscore +=15
} else if (sheet.getRange(lastrow,debate).getValue()=="To win at any cost"){
  careerscore +=25
}





Logger.log("careerscore");
Logger.log(careerscore);




careermsg = '';
careermsg = careermsg.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
careermsg = careermsg.concat('Your Career Path</p>');

if (careerscore > 120) {
careermsg = careermsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
careermsg = careermsg.concat('The Killer Shark<br></p>');
careermsg = careermsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
careermsg = careermsg.concat("You embody the burning spirit of youth, and tend to go all guns blazing no matter what the situation. Your fierce drive is your biggest strength, but can also turn into a weakness if not mitigated by a calming influence, an anchor to ground you in difficult times. Your neutral aggression is often mistaken for hostility, but your value to the organization is unmistakeable.");

} else if (careerscore > 65) {

careermsg = careermsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
careermsg = careermsg.concat('The Persistent Salmon<br></p>');
careermsg = careermsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
careermsg = careermsg.concat("You\'re a rare breed, not easily found in the seas of hiring, mostly because you\'ve already been taken for your tenacity at swimming upstream and taking on difficult challenges without any complaints. Your ability to disregard trivial things like the time of day and hunger when you\'re working makes you a formidable employee - just make sure you take care of your health (salmon are the most hunted fish after all).");

} else if (careerscore > 30) {

careermsg = careermsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
careermsg = careermsg.concat('The Carefree Dolphin<br></p>');
careermsg = careermsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
careermsg = careermsg.concat("You\'re smart and friendly - tend to go with the flow, and tend not to overthink things. Your career decisions are largely intuition based, without a lot of forethought - but they\'re usually pretty good since you have a general idea of where you want to go. Through these disjointed decisions, you are writing the story of your professional life, one seemingly impulsive move at a time.");

} else {

careermsg = careermsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
careermsg = careermsg.concat('The Confused Goldfish<br></p>');
careermsg = careermsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
careermsg = careermsg.concat("Unlike your fellow sharks or dolphins, you often fall prey to analysis paralysis - you overthink things and get too caught up in the \'what-ifs\' to actually take the decisions that need to be taken. But this is also your biggest strength in some cases, when your analysis finally finds the flaws that no one else can spot. But make sure you stop thinking more when you do - freeze in front of a disaster and you risk being eaten!");

}


careermsg = careermsg.concat('<br><hr></p>');
message = message.concat(careermsg);


//--------------------------------------------------------------------------------------------------------------------------------------
// Financial category
//--------------------------------------------------------------------------------------------------------------------------------------
var financialscore = 0

if (sheet.getRange(lastrow,maritalstatus).getValue() == "Married"){
  financialscore +=10
}
if (sheet.getRange(lastrow,gender).getValue() == "Male"){
  financialscore +=10
}
if (sheet.getRange(lastrow,partofday).getValue() == "Early morning" ){
  financialscore +=20
}
if (sheet.getRange(lastrow,subject).getValue() == "Maths/Science" ){
  financialscore +=20
}
if (sheet.getRange(lastrow,indoorsport).getValue().search("Strategy games (chess)")>=0 || sheet.getRange(lastrow,indoorsport).getValue().search("Card games (poker, bridge etc)")>=0 ){
  financialscore +=10
}
if (sheet.getRange(lastrow,demonetization).getValue() == "It was a terrible idea to begin with" ){
  financialscore +=15
}
if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I can't wait to get my hands on it" || sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll wait for the reviews"){
  financialscore +=15
}
if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll believe it when i see it"){
  financialscore +=5
}


if (sheet.getRange(lastrow,financialrisk).getValue() == "I invest in my present"){
  financialscore +=0
} else if (sheet.getRange(lastrow,financialrisk).getValue() == "Savings account" || sheet.getRange(lastrow,financialrisk).getValue() == "Insurance" ){
  financialscore +=20
} else if (sheet.getRange(lastrow,financialrisk).getValue() == "Mutual Funds"){
  financialscore +=30
} else if (sheet.getRange(lastrow,financialrisk).getValue() == "Short term stocks" || sheet.getRange(lastrow,financialrisk).getValue() == "Long term stocks"){
  financialscore +=50
} else if (sheet.getRange(lastrow,financialrisk).getValue() == "Real estate"){
  financialscore +=60
} 


Logger.log("financialscore");
Logger.log(financialscore);


finmsg = '';
finmsg = finmsg.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
finmsg = finmsg.concat('Financial Health</p>');

if (financialscore < 50) {

finmsg = finmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
finmsg = finmsg.concat('Maxed Out<br></p>');
finmsg = finmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
finmsg = finmsg.concat("If money travelled like Mumbai locals, you\'d be like Bhandup station, where the fast local doesn\'t really stop but just slows down for a bit before zooming away. You believe in living your present to its fullest rather than planning for your future, which seems too distant right now. But be warned - objects in the mirror are closer than they appear. Your past and present will catch up to your future in no time.");

} else if (financialscore < 90) {

finmsg = finmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
finmsg = finmsg.concat('The Wolf Of Wall Street<br></p>');
finmsg = finmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
finmsg = finmsg.concat("Most people think of money as a means to get what they want. But you know that the best use of money is to make more of it. You want to invest in everything - from low risk home loans and mutual funds to high frequency stocks. Just make sure you don\'t bite off more than you can chew.");

} else {

finmsg = finmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
finmsg = finmsg.concat('The Big Short<br></p>');
finmsg = finmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
finmsg = finmsg.concat("Money comes and money goes for most people. For you, the money stays in preparation for a possibly turbulent financial future. While most people are excited to spend their first salary, you patiently add most of it to your savings account. And though your long term strategy might be spot on, do try to lighten up and enjoy the present as and when you can.");

}


finmsg = finmsg.concat('<br><hr></p>');
message = message.concat(finmsg);



//--------------------------------------------------------------------------------------------------------------------------------------
// Lifestyle category - done
//--------------------------------------------------------------------------------------------------------------------------------------
var lifestylescore = 0


if (sheet.getRange(lastrow,agegroup).getValue() == "More than 40"){
  lifestylescore +=0
} else if (sheet.getRange(lastrow,agegroup).getValue() == "35-40"){
  lifestylescore +=5
} else if (sheet.getRange(lastrow,agegroup).getValue() == "30-35"){
  lifestylescore +=10
} else if (sheet.getRange(lastrow,agegroup).getValue() == "25-30"){
  lifestylescore +=15
} else if (sheet.getRange(lastrow,agegroup).getValue() == "20-25"){
  lifestylescore +=20
}

if (sheet.getRange(lastrow,maritalstatus).getValue() == "Single"){
  lifestylescore +=10
}
if (sheet.getRange(lastrow,gender).getValue() == "Male"){
  lifestylescore +=10
}
if (sheet.getRange(lastrow,tvtype).getValue().search("I don't watch TV")>=0 ){
  lifestylescore +=15
}


if (sheet.getRange(lastrow,vacation).getValue() == "Somewhere in the mountains" || sheet.getRange(lastrow,vacation).getValue() == "Wildlife reserve" ){
  lifestylescore +=20
} else if (sheet.getRange(lastrow,vacation).getValue() == "Somewhere sunny like a beach"){
  lifestylescore +=15
} else if (sheet.getRange(lastrow,vacation).getValue() == "A new city or country"){
  lifestylescore +=10
}


if (sheet.getRange(lastrow,partofday).getValue() == "Early morning"){
  lifestylescore +=30
} else if (sheet.getRange(lastrow,partofday).getValue() == "Later morning"){
  lifestylescore +=20
} else if (sheet.getRange(lastrow,partofday).getValue() == "Afternoon"){
  lifestylescore +=15
} else if (sheet.getRange(lastrow,partofday).getValue() == "Evening"){
  lifestylescore +=10
} else if (sheet.getRange(lastrow,partofday).getValue() == "Night"){
  lifestylescore +=5
} else if (sheet.getRange(lastrow,partofday).getValue() == "Late night" || sheet.getRange(lastrow,partofday).getValue() == "Whenever I'm awake!" ){
  lifestylescore +=0
} 

Logger.log("getting number of commas in string +1");
Logger.log(sheet.getRange(lastrow,physicalsport).getValue().split(",").length * 5)


lifestylescore += sheet.getRange(lastrow,physicalsport).getValue().split(",").length * 20
lifestylescore += sheet.getRange(lastrow,interestindiffsports).getValue().split(",").length * 7.5


Logger.log("lifestylescore");
Logger.log(lifestylescore);



lifestylemsg = '';
lifestylemsg = lifestylemsg.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
lifestylemsg = lifestylemsg.concat('Lifestyles</p>');


if (lifestylescore > 125 && sheet.getRange(lastrow,gender).getValue() == "Female") {

lifestylemsg = lifestylemsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lifestylemsg = lifestylemsg.concat('Xena, Warrior Princess<br></p>');
lifestylemsg = lifestylemsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lifestylemsg = lifestylemsg.concat(" Gymming and sports are your life outside of work. You have so many protein powder cans that your mom now uses them to fill her yearly supply of flour. You would rather miss a day at work than miss a day at gymming, and you have the body to show for it.");

} else if(lifestylescore > 125 && sheet.getRange(lastrow,gender).getValue() == "Female"){

lifestylemsg = lifestylemsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lifestylemsg = lifestylemsg.concat('The Arnold of Shivajinagar<br></p>');
lifestylemsg = lifestylemsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lifestylemsg = lifestylemsg.concat(" Gymming and sports are your life outside of work. You have so many protein powder cans that your mom now uses them to fill her yearly supply of flour. You would rather miss a day at work than miss a day at gymming, and you have the body to show for it.");

} else if(lifestylescore > 110){

lifestylemsg = lifestylemsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lifestylemsg = lifestylemsg.concat('The Human Calorimeter<br></p>');
lifestylemsg = lifestylemsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lifestylemsg = lifestylemsg.concat("You refuse sweets because they\'re high calorie, your yoga skills are second only to Baba Ramdev himself, and you satisfy your cravings for Indo-Chinese by whipping up a batch of Patanjali noodles at home rather than eat the a plate of greasy Manchurian. Amla juice is your preferred drink of choice, but even you take a break on birthdays and indulge in a slice of cake or two.");

} else if(lifestylescore > 75){

lifestylemsg = lifestylemsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lifestylemsg = lifestylemsg.concat('The Fearless Foodie<br></p>');
lifestylemsg = lifestylemsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lifestylemsg = lifestylemsg.concat("You like sports and fitness as much as anyone else, but that is no excuse for refusing a bowl of butter chicken or some dal makhani. You would rather gain a few kilos than lose the chance to sample some great desserts. But you don’\t like being mistaken for a glutton - you are a picky eater who only goes for the best (and 9 out of 10 times it\’s the food your mom makes).");

} else {

lifestylemsg = lifestylemsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
lifestylemsg = lifestylemsg.concat('The Club Invader<br></p>');
lifestylemsg = lifestylemsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
lifestylemsg = lifestylemsg.concat(" Your desire to party is never satisfied between Friday evening and Sunday night - a mere hangover isn\'t enough to stop you from ramping up the energy midweek and giving it another go - what are Redbull and Saridon for anyway? As for dark circles under your eyes - those are more like badges of honor than cause for genuine concern.");

}



lifestylemsg = lifestylemsg.concat('<br><hr></p>');
message = message.concat(lifestylemsg);



//--------------------------------------------------------------------------------------------------------------------------------------
// Temperament category - done
//--------------------------------------------------------------------------------------------------------------------------------------

var temperamentscore = 0

if (sheet.getRange(lastrow,numlanguages).getValue() == "1,2"){
  temperamentscore +=0
} else if (sheet.getRange(lastrow,numlanguages).getValue() == "3"){
  temperamentscore +=10
} else if (sheet.getRange(lastrow,numlanguages).getValue() == "4,5"){
  temperamentscore +=20
} else if (sheet.getRange(lastrow,numlanguages).getValue() == "> 5"){
  temperamentscore +=30
} 


if (sheet.getRange(lastrow,doctortype).getValue() == "I won't take medicines"){
  temperamentscore +=10
}


if (sheet.getRange(lastrow,mobileaddict).getValue().search("Banking apps")>=0 ){
  temperamentscore +=10
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Productivity apps")>=0 ){
  temperamentscore +=10
}
if (sheet.getRange(lastrow,mobileaddict).getValue().search("Educational apps")>=0 ){
  temperamentscore +=10
}
if (sheet.getRange(lastrow,partofday).getValue() == "Early morning"){
  temperamentscore +=15
} 
if (sheet.getRange(lastrow,tvtype).getValue() == "News"){
  temperamentscore +=10
} 


temperamentscore += sheet.getRange(lastrow,physicalsport).getValue().split(",").length * 10

if (sheet.getRange(lastrow,decisionlaziness).getValue() == "2010"){
  temperamentscore +=30
} else if (sheet.getRange(lastrow,decisionlaziness).getValue() == "2011"){
  temperamentscore +=25
} else if (sheet.getRange(lastrow,decisionlaziness).getValue() == "2012"){
  temperamentscore +=20
} else if (sheet.getRange(lastrow,decisionlaziness).getValue() == "2013-16"){
  temperamentscore +=10
} else if (sheet.getRange(lastrow,decisionlaziness).getValue() == "2017"){
  temperamentscore +=5
} 

if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I can't wait to get my hands on it"){
  temperamentscore +=15
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll wait for the reviews" || sheet.getRange(lastrow,ideadiffusion).getValue() == "I'll believe it when i see it"  ){
  temperamentscore +=10
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "My friends are likely to buy it before me"){
  temperamentscore +=5
} else if (sheet.getRange(lastrow,ideadiffusion).getValue() == "I'm pretty sure I won't buy it"){
  temperamentscore +=0
} 


if (sheet.getRange(lastrow,debate).getValue() == "To win at any cost"){
  temperamentscore +=10
} 

if (sheet.getRange(lastrow,internalexternal).getValue() == "I don't really know, things just happened" || sheet.getRange(lastrow,internalexternal).getValue() == "My parents/family/better half asked me to"){
  temperamentscore +=0
} else if (sheet.getRange(lastrow,internalexternal).getValue() == "I like the money" || sheet.getRange(lastrow,internalexternal).getValue() == "I don't know where I want to go, and this profession buys me time to think about it"){
  temperamentscore +=10
} else if (sheet.getRange(lastrow,internalexternal).getValue() == "What I do at work allows me to do what I want outside work"){
  temperamentscore +=15
} else if (sheet.getRange(lastrow,internalexternal).getValue() == "I know where I want to go and this profession allows me to get there"){
  temperamentscore +=25
} 


Logger.log("temperamentscore");
Logger.log(temperamentscore);


temperamentmsg = '';
temperamentmsg = temperamentmsg.concat('<p style="font-size: 30px;font-family: Palatino, serif;color:663399;text-decoration: underline;">');
temperamentmsg = temperamentmsg.concat('Your Attitudes</p>');

if (temperamentscore > 125) {

temperamentmsg = temperamentmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
temperamentmsg = temperamentmsg.concat('The Hyper Viper<br></p>');
temperamentmsg = temperamentmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
temperamentmsg = temperamentmsg.concat(" You are dangerously active and are excited by the tiniest of incidents. People around you are afraid of your impulsiveness and impatience, but you know that it\'s all under control, and for a reason. It\'s the minute details that matter to you and you leave no stone unturned if they\'re not right. You are the kind that takes exact 22 yards measurements with their feet even while playing gully cricket, but you also hit the most sixes!");

} else if (temperamentscore > 75) {

temperamentmsg = temperamentmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
temperamentmsg = temperamentmsg.concat('The Energetic Iguana<br></p>');
temperamentmsg = temperamentmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
temperamentmsg = temperamentmsg.concat("You\'re an active person and probably have done quite a few extra-curricular activities in your school days. These days you apply the same spirit to your work and it\'s hard to break you. You know exactly where the goal is and the right platforms to showcase your talent. You are someone who gets the job done in time without any advese consequences.");

} else if (temperamentscore > 35) {

temperamentmsg = temperamentmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
temperamentmsg = temperamentmsg.concat('The Alert Alligator<br></p>');
temperamentmsg = temperamentmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
temperamentmsg = temperamentmsg.concat("Procrastination is your favourite music. You like to relax a lot (in the sun) and it helps you stay in a good mood, but you\’re also alert for spotting your prey! When you\'re active, you like doing a lot of things simultaneously, but rarely finish them all. In spite of your generally chilled outlook, you know your priorities well and do care about your work, but sometimes the call of duty on your PS3 is too strong to ignore!!");

} else {

temperamentmsg = temperamentmsg.concat('<p style="font-size: 22px;font-family: Palatino, serif;font-weight: 700;color:444444;">');
temperamentmsg = temperamentmsg.concat('The Sloth<br></p>');
temperamentmsg = temperamentmsg.concat('<p style="font-size: 19px;font-family: Palatino, serif;color:444444;">');
temperamentmsg = temperamentmsg.concat("The sun barely moves, if at all, and the day is fifty hours long. You look constantly out the window, to gaze carefully at the sun to determine how far it stands from lunchtime. That basically defines your typical daily decision-making. The less work the better. You have found all the shortcuts in the world, and you would call your roommate even if he is in the next room. If there was an award for laziness you would probably send someone to pick it up for you.");

}


temperamentmsg = temperamentmsg.concat('<br><hr></p>');
message = message.concat(temperamentmsg);

//--------------------------------------------------------------------------------------------------------------------------------------
// Share Buttons
//--------------------------------------------------------------------------------------------------------------------------------------
sharescript = '';
sharescript = sharescript.concat('<p style="font-size: 16px;font-family: Palatino, serif;color: #8C001A;">');
sharescript = sharescript.concat('<br><i><b>So you liked the quiz: Why not share it?</i></b><br></p>');







sharescript = sharescript.concat('<a href="https://api.addthis.com/oexchange/0.8/forward/facebook/offer?url=https%3A%2F%2Fgoo.gl%2Fforms%2FImhuWGgHJ743c5c82&pubid=ra-42ankitrahul42&title=The%20Nirvana%20Game&ct=1" target="_blank"><img src="cid:facebookImage" border="0" alt="Facebook"/></a>');
sharescript = sharescript.concat('<a href="https://api.addthis.com/oexchange/0.8/forward/google_plusone_share/offer?url=https%3A%2F%2Fgoo.gl%2Fforms%2FImhuWGgHJ743c5c82&pubid=ra-42ankitrahul42&title=The%20Nirvana%20Game&ct=1" target="_blank"><img src="cid:gplusImage" border="0" alt="Google+"/></a>');
sharescript = sharescript.concat('<a href="https://api.addthis.com/oexchange/0.8/forward/twitter/offer?url=https%3A%2F%2Fgoo.gl%2Fforms%2FImhuWGgHJ743c5c82&pubid=ra-42ankitrahul42&title=The%20Nirvana%20Game&ct=1" target="_blank"><img src="cid:twitterImage" border="0" alt="Twitter"/></a>');
sharescript = sharescript.concat('<a href="https://api.addthis.com/oexchange/0.8/forward/whatsapp/offer?url=https%3A%2F%2Fgoo.gl%2Fforms%2FImhuWGgHJ743c5c82&pubid=ra-42ankitrahul42&title=The%20Nirvana%20Game&ct=1" target="_blank"><img src="cid:whatsappImage" border="0" alt="WhatsApp"/></a>');
sharescript = sharescript.concat('<a href="https://api.addthis.com/oexchange/0.8/forward/email/offer?url=https%3A%2F%2Fgoo.gl%2Fforms%2FImhuWGgHJ743c5c82&pubid=ra-42ankitrahul42&title=The%20Nirvana%20Game&ct=1" target="_blank"><img src="cid:emailImage" border="0" alt="Email"/></a>');

sharescript = sharescript.concat('<br>');
message = message.concat(sharescript);

//--------------------------------------------------------------------------------------------------------------------------------------
// Disclaimer
//--------------------------------------------------------------------------------------------------------------------------------------
disclaimer = '';
disclaimer = disclaimer.concat('<p style="font-size: 12px;font-family: Palatino, serif;color:#333333;">');
disclaimer = disclaimer.concat('*Disclaimer: The survey results are in the first stage of development and may not accurately reflect your personality. You accurate answers help us build our model better. Thank you for taking the survey</p>');






disclaimer = disclaimer.concat('</p>');
message = message.concat(disclaimer);

//--------------------------------------------------------------------------------------------------------------------------------------
// Send the messages and set confirmation
//--------------------------------------------------------------------------------------------------------------------------------------
                          
   MailApp.sendEmail({
     to: email,
     bcc:"rahul.madhavan@atidiv.com,ankit.baraskar@atidiv.com",
     subject: mailSubject,
     htmlBody: message,
     inlineImages:
       {
         personalityImage: imageBlob,
         facebookImage: facebookBlob,
         whatsappImage: whatsappBlob,
         emailImage: emailBlob,
         gplusImage: gplusBlob,
         twitterImage: twitterBlob,
       }
   });
   
   
   
   }
   
   
   
      // Returns a confirmation message whether email was sent, in the Column designated in step 1
      // You can change these confirmation messages, or get rid of them altogether by making these
      // lines comment lines by adding the "//" to the beginning of the next 3 lines    
      
      sheet.getRange(lastrow,columnnumber,1,1).setValue("Email Sent");
      }  
     else{
         sheet.getRange(lastrow,columnnumber,1,1).setValue("Email not Sent - Invalid Email Address submitted");
      }   
Logger.log(Date.now() )
} 

'created by Rui Santos, http://randomnerdtutorials.wordpress.com, 2013
'modified by Laura Boccanfuso, Lisa Chen and Colette Torres 6/2015
'Control a servo motor with Visual Basic 
Imports System.Globalization
Imports System.IO
Imports System.IO.Ports
Imports System.Security.Cryptography.X509Certificates
Imports System.Speech.Synthesis
Imports System.Text
Imports System.Threading
Imports Fleck
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Shared _continue As Boolean
    Shared _serialPort As SerialPort
    WithEvents speaker As New SpeechSynthesizer()
    Public Event VisemeReached As EventHandler(Of VisemeReachedEventArgs)
    Public Event SpeakCompleted As EventHandler(Of SpeakCompletedEventArgs)
    Dim stop_Clicked As Boolean = False
    Dim pause_Clicked As Boolean = False
    Dim myName As String
    Dim failedToConnect As Boolean
    ReadOnly LEVoice As String = "IVONA 2 Ivy OEM" 'Microsoft Anna OR IVONA 2 Ivy OEM
    Dim currentThread As Thread

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '#JC Init the Speech Recognition
        InitSpeechRecoginition()

        'Setup Voice
        speaker.Rate = -2
        speaker.Volume = 100
        Try
            speaker.SelectVoice(LEVoice)
        Catch ex As Exception

        End Try

        SerialPort1.Close()
        SerialPort1.PortName = "COM6" 'define your port
        SerialPort1.BaudRate = 9600
        SerialPort1.DataBits = 8
        SerialPort1.Parity = Parity.None
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Handshake = Handshake.None
        SerialPort1.Encoding = Encoding.Default
        Thread.Sleep(1000)
        Try
            SendCommand("2")
        Catch ex As Exception
            If (ex.ToString.Contains("does not exist")) Then
                MsgBox(
                    "Failed to connect to serial port. Please make sure L-E is plugged in and the arduino code has been uploaded.")
                failedToConnect = True
            End If
        End Try
        EnableTimers()
    End Sub

    Private Sub speaker_VisemeReached(sender As Object, e2 As VisemeReachedEventArgs) Handles speaker.VisemeReached
        'Console.WriteLine("Viseme " & e2.Viseme & " was " & e2.Duration.TotalMilliseconds.ToString & " ms. long" & vbNewLine)

        '0:      silence()
        '1:      ae, ax, ah
        '2:      aa()
        '3:      ao()
        '4:      ey, eh, uh
        '5:      er()
        '6:      y, iy, ih, ix
        '7:      w, uw4
        '8:      ow()
        '9:      aw()
        '10:     oy()
        '11:     ay()
        '12:     h()
        '13:     r()
        '14:     l()
        '15:     s, z
        '16:     sh, ch, jh, zh
        '17:     th, dh
        '18:     f, v
        '19:     d, t, n
        '20:     k, g, ng
        '21:     p, b, m
        Try
            If (e2.Viseme = 1 Or e2.Viseme = 2 Or e2.Viseme = 9 Or e2.Viseme = 8) Then
                SendCommand("4")
            ElseIf (e2.Viseme = 3 Or e2.Viseme = 6 Or e2.Viseme = 10) Then
                SendCommand("5")
            ElseIf (e2.Viseme = 4 Or e2.Viseme = 5 Or e2.Viseme = 7 Or e2.Viseme = 20) Then
                SendCommand("6")
            Else
                SendCommand("7")
            End If
        Catch e As Exception
            'Console.WriteLine("exception caught")
        End Try
    End Sub

    Private Sub speaker_SpeakCompleted(sender As Object, e2 As SpeakCompletedEventArgs) Handles speaker.SpeakCompleted
        If (stop_Clicked = True) Then
            Return
        Else
            EnableTimers()
        End If
    End Sub

    Private Sub NameBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NameBox.KeyPress
        DisableTimers()
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            myName = NameBox.Text
        End If
        EnableTimers()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        DisableTimers()
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SpeakString(TextBox1.Text)
        End If
        EnableTimers()
    End Sub

    Private Sub Clear_Click(sender As Object, e As EventArgs) Handles Clear.Click
        TextBox1.Clear()
    End Sub

#Region "Faces"

    Private Sub neutral_Click(sender As Object, e As EventArgs) Handles Timer3.Tick, neutral.Click
        SendCommand("2")
    End Sub

    Private Sub smile_Click_1(sender As Object, e As EventArgs) Handles smile.Click
        SendCommand("0", True)
    End Sub

    Private Sub sad_Click(sender As Object, e As EventArgs) Handles frown.Click
        SendCommand("1", True)
    End Sub

    Private Sub Confused_Click(sender As Object, e As EventArgs) Handles Confused.Click
        SendCommand("F", True)
    End Sub

    Private Sub Surprised_Click(sender As Object, e As EventArgs) Handles Surprised.Click
        SendCommand("G", True)
    End Sub

    Private Sub Angry_Click(sender As Object, e As EventArgs) Handles Angry.Click
        SendCommand("H", True)
    End Sub

    Private Sub CrossEyed_Click(sender As Object, e As EventArgs) Handles CrossEyed.Click
        SendCommand("I", True)
    End Sub

    Private Sub Awkward_Click(sender As Object, e As EventArgs) Handles Awkward.Click
        SendCommand("J", True)
    End Sub

    Private Sub FunnyFace1_Click(sender As Object, e As EventArgs) Handles FunnyFace1.Click
        SendCommand("K", True)
    End Sub

    Private Sub Afraid_Click(sender As Object, e As EventArgs) Handles Afraid.Click
        SendCommand("L", True)
    End Sub

    Private Sub Sleepy_Click(sender As Object, e As EventArgs) Handles Sleepy.Click
        SendCommand("M", True)
    End Sub

    Private Sub Yelling_Click(sender As Object, e As EventArgs) Handles Yelling.Click
        SendCommand("N", True)
    End Sub

    Private Sub Animated_Click(sender As Object, e As EventArgs) Handles Animated.Click
        SendCommand("R", True)
    End Sub

    Private Sub Funny2_Click(sender As Object, e As EventArgs) Handles Funny2.Click
        SendCommand("Q")
    End Sub

#End Region

#Region "Actions/Body Positions"

    Private Sub blink_Click(sender As Object, e As EventArgs) Handles Timer1.Tick, blink.Click
        SendCommand("3")
    End Sub

    Private Sub EyeRight_Click(sender As Object, e As EventArgs) Handles EyeRight.Click
        SendCommand("8", True)
        EyeRight.BackColor = Color.Gold
    End Sub

    Private Sub EyeLeft_Click(sender As Object, e As EventArgs) Handles EyeLeft.Click
        SendCommand("9", True)
        EyeLeft.BackColor = Color.Gold
    End Sub

    Private Sub HeadLeft_Click(sender As Object, e As EventArgs) Handles Timer2.Tick, HeadLeft.Click
        SendCommand("A")
        HeadLeft.BackColor = Color.Gold
    End Sub

    Private Sub HeadRight_Click(sender As Object, e As EventArgs) Handles HeadRight.Click
        SendCommand("B", True)
        HeadRight.BackColor = Color.Gold
    End Sub

    Private Sub HeadUp_Click(sender As Object, e As EventArgs) Handles HeadUp.Click
        SendCommand("C", True)
        HeadUp.BackColor = Color.Gold
    End Sub

    Private Sub HeadDown_Click(sender As Object, e As EventArgs) Handles HeadDown.Click
        SendCommand("D", True)
        HeadUp.BackColor = Color.Gold
    End Sub

    Private Sub Wink_Click(sender As Object, e As EventArgs) Handles Wink.Click
        SendCommand("E", True)
    End Sub

#End Region

#Region "One-Worders"

    Private Sub Oh_Click(sender As Object, e As EventArgs) Handles Oh.Click
        SpeakString("Oh", -2, 100, True)
        Oh.BackColor = Color.Gold
    End Sub

    Private Sub Okay_Click(sender As Object, e As EventArgs) Handles Okay.Click
        SpeakString("Okay...", -2, 100, True)
        Okay.BackColor = Color.Gold
    End Sub

    Private Sub Wow_Click(sender As Object, e As EventArgs) Handles Wow.Click
        SpeakString("Wow.", -2, 100, True)
        Wow.BackColor = Color.Gold
    End Sub

    Private Sub Interesting_Click(sender As Object, e As EventArgs) Handles Interesting.Click
        SpeakString("Interesting.....", -2, 100, True)
        EnableTimers()
        Interesting.BackColor = Color.Gold
    End Sub

    Private Sub Cool_Click(sender As Object, e As EventArgs) Handles Cool.Click
        SpeakString("Cool", -2, 100, True)
        Cool.BackColor = Color.Gold
    End Sub

    Private Sub Nice_Click(sender As Object, e As EventArgs) Handles Nice.Click
        SpeakString("Nice", -2, 100, True)
        Nice.BackColor = Color.Gold
    End Sub

    Private Sub Yeah_Click(sender As Object, e As EventArgs) Handles Yeah.Click
        SpeakString("Yeah", -2, 100, True)
        Yeah.BackColor = Color.Gold
    End Sub

    Private Sub Uhuh_Click(sender As Object, e As EventArgs) Handles Uhuh.Click
        SpeakString("Uhhhhh huh....", -2, 100, True)
        Uhuh.BackColor = Color.Gold
    End Sub

#End Region

#Region "Small Talk"

    Private Sub HowAreYou_Click(sender As Object, e As EventArgs) Handles HowAreYou.Click
        SpeakString("How are you?", -2, 100, True)
    End Sub

    Private Sub ImGood_Click(sender As Object, e As EventArgs) Handles ImGood.Click
        SpeakString("I'm good", -2, 100, True)
    End Sub

    Private Sub ImOkay_Click(sender As Object, e As EventArgs) Handles ImOkay.Click
        SpeakString("I'm okay", -2, 100, True)
    End Sub

    Private Sub NotGreat_Click(sender As Object, e As EventArgs) Handles NotGreat.Click
        SpeakString("Not great.", -2, 100, True)
    End Sub

    Private Sub NotGreatCont_Click(sender As Object, e As EventArgs) Handles NotGreatCont.Click
        SpeakString("I had too much math homework and I couldn't figure out a bunch of the problems", -2, 100, True)
    End Sub

    Private Sub TodayQ_Click(sender As Object, e As EventArgs) Handles TodayQ.Click
        SpeakString("What did you do today?", -2, 100, True)
    End Sub

    Private Sub TodayA_Click(sender As Object, e As EventArgs) Handles TodayA.Click
        SpeakString("Today, I went to the park", -2, 100, True)
    End Sub

    Private Sub TodayACont_Click(sender As Object, e As EventArgs) Handles TodayACont.Click
        SpeakString("There was some really cool music playing in the park", -2, 100, True)
    End Sub
    Private Sub AboutYrslf_Click(sender As Object, e As EventArgs) Handles AboutYrslf.Click
        SpeakString("Tell me about yourself!")
    End Sub
#End Region

#Region "Scripts"

    Sub Thread1Task()
        Dim string2say = "Hello. I’m good."
        Dim string2say2 = "I went to my friend’s birthday party."
        Dim string2say3 = "Yeah."
        Dim string2say4 = "I went to a pool party and I swam and I ate pizza and cake."
        Dim string2say5 = "Yeah. I played on the playground. The slide was very big."
        Dim string2say6 = "Chocolate."
        Dim string2say7 = "Okay."
        Dim string2say8 = "Oh yeah, and I also played with my friend Jimmy. We went off the diving board. I did a cannon ball."
        Dim string2say9 = "Okay."
        Dim string2say10 = "That's nice"
        Dim string2say11 = "Oh yeah. And this weekend, I saw a really cool car."
        Dim string2say12 = "Yes."

        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(2000)          'Hi L-E. How are you?
        SendCommand("0")            'smiles
        SpeakString(string2say)     'Hello. I'm good.
        Thread.Sleep(1700)          'What did you do this weekend? 
        SpeakString(string2say2)    'I went to my friend’s birthday party.
        Thread.Sleep(3250)          'Oh yeah? Was it fun? What did you do?
        SendCommand("3")            'blinks
        SpeakString(string2say3)    'Yeah
        Thread.Sleep(6000)          'So what did you do at the party? (long pause)... L-E?
        SpeakString(string2say4)    'I went to a pool party and I swam and I ate pizza and cake.
        Thread.Sleep(2750)          'That’s cool, what’s your favorite kind of cake?
        SpeakString(string2say5)    'Yeah, I played on the playground. The slide was very big.
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(3000)          'Oh, but what’s your favorite kind of cake?
        SendCommand("3")            'blinks
        Thread.Sleep(500)           'pause to finish blinking
        SpeakString(string2say6)    'Chocolate.
        Thread.Sleep(2000)          'Can I tell you about my weekend?
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(1000)          'pause to finish turning
        SpeakString(string2say7)    'Okay
        Thread.Sleep(500)           'I went to the--
        SpeakString(string2say8)    'Oh yeah, and I also played with my friend Jimmy. We went off the diving board. I did a cannon ball.
        Thread.Sleep(4000)          'That’s very nice, L-E. Can I finish telling you about my weekend?
        SpeakString(string2say9)    'Okay
        Thread.Sleep(4250)          'I went to the beach, and I fed seagulls and found some really pretty seashells.
        SendCommand("3")            'blinks
        SendCommand("B")            'looks away
        SendCommand("8")            'looks away
        Thread.Sleep(800)           'pause for motors to finish
        SpeakString(string2say10)   'That's nice 
        Thread.Sleep(1800)          'Yeah. Do you like going to the beach?
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(500)           'pause to finish turning
        SpeakString(string2say11)   'Oh yeah. And this weekend, I saw a really cool car.
        Thread.Sleep(4250)          'Oh, that’s fun.  But I was asking if you like the beach…
        SpeakString(string2say12)   'Yes
        Thread.Sleep(3000)          'Great. Maybe one day, we can take a trip to the beach together!
    End Sub

    Private Sub Script1_Click(sender As Object, e As EventArgs) Handles Script1.Click
        currentThread = New Thread(AddressOf Thread1Task)
        DisableTimers()
        currentThread.Start()
    End Sub

    Private Sub Thread2Task()
        Dim string2say = "I like to play Super Mario Bros."
        Dim string2say2 = "I also like this cool Lego game."
        Dim string2say3 = "Luigi."
        Dim string2say4 = "I like him."
        Dim string2say5 = "My favorite color is green."
        Dim string2say6 = "Oh yeah, in the Lego Game my character is also green."
        Dim string2say7 = "Wow, we have the same favorite color."
        Dim string2say8 = "I like watching TV"
        Dim string2say9 = "Phineas and Ferb"
        Dim string2say10 = "Yeah"
        Dim string2say11 = "Summer is the best time of the year."
        Dim string2say12 = "Oh yeah.  I also really like the movie Despicable Me"
        Dim string2say13 = "I like to play with my toys"

        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(3000)          'What’s your favorite video game, L-E?
        SpeakString(string2say)     'I like to play Super Mario Bros.
        Thread.Sleep(2500)          'How cool!  Who’s your favorite character?
        SpeakString(string2say2)    'I also like this cool Lego game.
        Thread.Sleep(3500)          'That’s nice, but who’s your favorite character in Mario Bros?
        SpeakString(string2say3)    'Luigi.
        Thread.Sleep(2750)          'I see.  And why is he your favorite?
        SpeakString(string2say4)    'I like him.
        Thread.Sleep(2000)          'Yes, why do you like him?
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(1000)          'pause for motors 
        SpeakString(string2say5)    'My favorite color is green.
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(1150)          'Awesome.  My favorite color--
        SpeakString(string2say6)    'Oh yeah, in the Lego Game my character is also green.
        Thread.Sleep(7000)          'That’s nice, L-E.  I was just saying that my favorite color is also green... (long pause)
        SendCommand("3")            'blink eyes
        Thread.Sleep(600)           'pause for motors 
        SpeakString(string2say7)    'Wow, we have the same favorite color.
        Thread.Sleep(4250)          'Yeah! So what else do you like to do when you’re not playing video games?
        SendCommand("B")            'looks away
        SendCommand("8")
        SpeakString(string2say8)    'I like watching TV.
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(3000)          'Cool!  What's your favorite TV show? 
        SpeakString(string2say9)    'Phineas and Ferb.
        Thread.Sleep(1750)          'Oh yeah? Why is that?           
        SpeakString(string2say10)   'Yeah.
        Thread.Sleep(2000)          'Why do you like that show, L-E?
        SpeakString(string2say11)   'Summer is the best time of the year.
        Thread.Sleep(4250)          'I agree! Do you ever do cool things like Phineas and Ferb over the summer?
        SpeakString(string2say12)   'Oh yeah.  I also really like the movie “Despicable Me”
        Thread.Sleep(6000)          'That’s nice, L-E.  I just saw that movie too, but I was just asking you what you do in the summer…
        SpeakString(string2say13)   'I like to play with my toys.
        Thread.Sleep(3000)          'How nice! Toys are always fun!
    End Sub

    Private Sub Script2_Click(sender As Object, e As EventArgs) Handles Script2.Click
        currentThread = New Thread(AddressOf Thread2Task)
        DisableTimers()
        currentThread.Start()
    End Sub

    Private Sub Thread3Task()
        Dim string2say = " School is fun sometimes."
        Dim string2say2 = "I like recess. "
        Dim string2say3 = "I hate math."
        Dim string2say4 = "Uh yeah."
        Dim string2say5 = "Math is boring."
        Dim string2say6 = "Yeah, well... I don't like doing math homework"
        Dim string2say7 = "Yeah"
        Dim string2say8 = "I like her a lot. She’s my favorite teacher so far."
        Dim string2say9 = "Mrs. Peterson. She gives me a lot of stickers."
        Dim string2say10 = "Once, Mrs. Peterson made cookies for the class."
        Dim string2say11 = " I like the scratch and sniff stickers."
        Dim string2say12 = "Do you love cookies? I love cookies!"
        Dim string2say13 = "Does that mean we can have some now?"

        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(1500)          'So how do you like school, L-E?
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(1000)          'pause for motors 
        SpeakString(string2say)     'School is fun sometimes.
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(2000)          'Oh yeah?  What do you like to do at school?
        SpeakString(string2say2)    'I like recess.
        Thread.Sleep(1000)          'I love recess too. What's your fav--
        SpeakString(string2say3)    'I hate math.
        Thread.Sleep(2000)          'Why do you hate math?
        speaker.Rate = -8
        SpeakString(string2say4)    'Uhh yeah.
        speaker.Rate = -2
        SendCommand("3")            'blink
        Thread.Sleep(1000)          'L-E? Why don't you like math?
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(1000)          'pause for motors 
        SpeakString(string2say5)    'Math is boring.
        Thread.Sleep(1000)          'stay turned for a second
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(5750)          'Well that’s too bad.  I really liked math... (long pause) 
        SpeakString(string2say6)    'Yeah, well, I don't like doing math homework.
        Thread.Sleep(2000)          'Well how do you like your teacher?
        SendCommand("3")            'blink
        Thread.Sleep(600)           'pause for motors 
        SpeakString(string2say7)    'Yeah
        Thread.Sleep(4750)          'Yeah, you like her? Can you tell me a bit more about your teacher?
        SpeakString(string2say8)    'I like her a lot. She’s my favorite teacher so far.
        Thread.Sleep(2600)          'That's great.  What's her name?
        SpeakString(string2say9)    'Mrs. Peterson. She gives me a lot of stickers.
        Thread.Sleep(2500)          'Oh! What kind of stickers?
        SpeakString(string2say10)   'Once, Mrs. Peterson made cookies for the class.
        Thread.Sleep(5750)          'That’s really nice of Mrs. Peterson, but I was just asking you about what kind of stickers you like?
        SpeakString(string2say11)   'I like the scratch and sniff stickers
        Thread.Sleep(3000)          'Those are fun stickers! I love animal stickers.
        SpeakString(string2say12)   'Do you love cookies?  I love cookies!
        Thread.Sleep(2500)          'Oh! Uh... me too!
        SpeakString(string2say13)   'Does that mean we can have some now?
        SendCommand("E")            'wink
        Thread.Sleep(3000)          'No, I don't think so, L-E...  Maybe next time! 
    End Sub

    Private Sub Script3_Click(sender As Object, e As EventArgs) Handles Script3.Click
        currentThread = New Thread(AddressOf Thread3Task)
        DisableTimers()
        currentThread.Start()
    End Sub

    Private Sub Thread4Task()
        Dim string2say As String = "Yes. I would like to be a doctor one day."
        Dim string2say2 As String = "I’ve wanted to be a doctor ever since I got my first toy doctor kit"
        Dim string2say3 As String = "I got it when I was 5"
        Dim string2say4 As String = "It has a stethoscope, thermometer, bandages...  Stuff like that"
        Dim string2say5 As String = "I really like helping people and making them feel better."
        Dim string2say6 As String = "I have a friend who is a doctor. She’s awesome."
        Dim string2say7 As String = "A heart doctor."
        Dim string2say8 As String = "Yeah."
        Dim string2say9 As String = "A children's doctor"
        Dim string2say10 As String = "Yeah."
        Dim string2say11 As String = "Because children are really cool."
        Dim string2say12 As String = "My friend has a cool model of a heart in her office, and once she let me play with it."
        Dim string2say13 As String = "Yeah, he gives me lollipops."
        Dim string2say14 As String = "Yeah. My favorite flavor is blue raspberry"

        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(5000)          'So, L-E, do you have any idea what you want to be when you grow up?
        SpeakString(string2say)     'Yes. I would like to be a doctor one day.
        Thread.Sleep(900)           'That's awesome. Why do you wan--
        SpeakString(string2say2)    'I’ve wanted to be a doctor ever since I got my first toy doctor kit.
        Thread.Sleep(3000)          'Oh yeah? How old were you when you got it?
        SpeakString(string2say3)    'I got it when I was five.
        Thread.Sleep(6000)          'Nice... I've never seen a toy doctor kit.  Can you tell me more about what it has?
        SpeakString(string2say4)    'It has a stethoscope, thermometer, bandages... Stuff like that.
        Thread.Sleep(2850)          'That's so cool! Who got you that doctor kit?
        SpeakString(string2say5)    'I really like helping people and making them feel better.
        Thread.Sleep(5500)          'Yeah that’s nice, but I was wondering who got you that doctor kit... Do you know any doctors?
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(1000)          'pause for motors
        SpeakString(string2say6)    'I have a friend who is a doctor. She’s awesome.
        Thread.Sleep(1000)          'pause in that position for a sec 
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(1500)          'That’s great! What kind of doctor is your friend?
        SpeakString(string2say7)    'A heart doctor.
        Thread.Sleep(3500)          'Woah, cool! What kind of doctor do you want to be?
        SpeakString(string2say8)    'Yeah.
        Thread.Sleep(2500)          'Soo... What kind of doctor do you want to be?
        SpeakString(string2say9)    'A children’s doctor. 
        Thread.Sleep(2000)          'Oh wow! Why is that?
        SpeakString(string2say10)   'Yeah.
        Thread.Sleep(1200)          'Yeah, why?
        SendCommand("B")            'looks away
        SendCommand("8")
        SpeakString(string2say11)  'Because children are really cool.
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(5000)          'That is true! Children's doctors are also cool. Do you like your doctor?
        SpeakString(string2say12)   'My friend has a cool model of a heart in her office, and once, she let me play with it!
        Thread.Sleep(4000)          'That sounds pretty neat, but I was just asking you if you liked your doctor.
        SendCommand("3")            'blink
        Thread.Sleep(3000)          'long pause
        SpeakString(string2say13)   'Yeah. He gives me lollipops. 
        Thread.Sleep(2000)          'Nice! I love lollipops! 
        SpeakString(string2say14)   'Yeah. My favorite flavor is blue raspberry!
        Thread.Sleep(2000)          'Oh! Me too!
    End Sub

    Private Sub Script4_Click(sender As Object, e As EventArgs) Handles Script4.Click
        currentThread = New Thread(AddressOf Thread4Task)
        DisableTimers()
        currentThread.Start()
    End Sub

    Sub Thread5Task()
        Dim string2say = "Hello."
        Dim string2say2 = "I hate roller coasters"
        Dim string2say3 = "I like chocolate cake"
        Dim string2say4 = "I especially like chocolate cake with peanut butter frosting"
        Dim string2say5 = "Yeah. Peanut butter is the best!"
        Dim string2say6 = "Yeah"
        Dim string2say7 = "Yes, PB&J sandwiches are my favorite kind of lunch"
        Dim string2say8 = "And Peanut butter cups are my favorite chocolate"
        Dim string2say9 = "That's interesting"

        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(1000)          'Hi L-E. 
        SpeakString(string2say)     'Hello
        Thread.Sleep(2000)          'What's your favorite kind of food?
        SpeakString(string2say2)    'I hate roller coasters
        Thread.Sleep(4250)          'Oh, I don't like them either, but I was asking about your favorite kind of food
        SpeakString(string2say3)    'I like chocolate cake
        Thread.Sleep(1750)          'Oh, wow! Me too!
        SendCommand("B")            'looks away
        SendCommand("8")
        Thread.Sleep(1000)          'pause for motors
        SpeakString(string2say4)    'I espeecially like chocolate cake with peanut butter frosting
        Thread.Sleep(1000)          'pause in that position for a sec
        SendCommand("A")            'turns to speaker
        SendCommand("9")
        Thread.Sleep(2000)          'That sounds really good.  Do you love peanut butter?
        SpeakString(string2say5)    'Yeah, peanut butter is the best!
        Thread.Sleep(3000)          'I see! What do you like eating peanut butter with?
        SpeakString(string2say6)    'Yeah.
        Thread.Sleep(6500)          'So do you eat peanut butter in sandwiches?... (long pause)
        SpeakString(string2say7)    'Yes, PB&J sandwiches are my favorite kind of lunch
        Thread.Sleep(1000)          'Nice. I like--
        SpeakString(string2say8)    'Peanut butter cups are my favorite chocolate!
        Thread.Sleep(5000)          'That's nice, L-E. I was just saying that I like to eat peanut butter sandwiches with bananas.
        SpeakString(string2say9)    'That's interesting
        Thread.Sleep(500)           'Yup!
    End Sub

    Private Sub Demo_Click(sender As Object, e As EventArgs) Handles Demo.Click
        currentThread = New Thread(AddressOf Thread5Task)
        DisableTimers()
        currentThread.Start()
    End Sub

#End Region

#Region "Stallers"

    Private Sub Staller1_Click(sender As Object, e As EventArgs) Handles Staller1.Click
        SpeakString("That's a good question!... Let me think...", 0.2, 100, True)
    End Sub

    Private Sub Staller2_Click(sender As Object, e As EventArgs) Handles Staller2.Click
        SpeakString("I'm not sure...  I'll have to ask my mom.", 0.2, 100, True)
    End Sub

    Private Sub Staller3_Click(sender As Object, e As EventArgs) Handles Staller3.Click
        SpeakString("Welllllll....", 0.5, 100, True)
    End Sub

    Private Sub Staller4_Click(sender As Object, e As EventArgs) Handles Staller4.Click
        SpeakString("Sorry, can you say that again?", 0.2, 100, True)
    End Sub

#End Region

#Region "Response Timing"

    Private Sub Interrupt1_Click(sender As Object, e As EventArgs) Handles Interrupt1.Click
        DisableTimers()
        SendCommand("R")
        SpeakString("Oh yeah and I forgot that I wanted to say something else")
        Interrupt1.BackColor = Color.Gold
        EnableTimers()
    End Sub

    Private Sub Interrupt2_Click(sender As Object, e As EventArgs) Handles Interrupt2.Click
        DisableTimers()
        SendCommand("R")
        SpeakString("Also there was something else I wanted to add...")
        Interrupt2.BackColor = Color.Gold
        EnableTimers()
    End Sub

    Private Sub DelayBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DelayBox.KeyPress
        DisableTimers()
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Thread.Sleep(2500)
            SpeakString(DelayBox.Text, 0.2)
            DelayBox.Clear()
            DelayBox.BackColor = Color.Gold
        End If
        EnableTimers()
    End Sub

    Private Sub DelayClear_Click(sender As Object, e As EventArgs) Handles DelayClear.Click
        DelayBox.Clear()
    End Sub

#End Region

#Region "Common Convo"

    Private Sub LikeBox_KeyDown(sender As Object, e As KeyPressEventArgs) Handles LikeBox.KeyPress
        DisableTimers()
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SpeakString("I like " + LikeBox.Text, 0.2)
            LikeBox.Clear()
        End If
        EnableTimers()
    End Sub

    Private Sub DontLikeBox_KeyDown(sender As Object, e As KeyPressEventArgs) Handles DontLikeBox.KeyPress
        DisableTimers()
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SpeakString("I don't like " + DontLikeBox.Text, 0.2)
            DontLikeBox.Clear()
        End If
        EnableTimers()
    End Sub

    Private Sub LikeClear_Click(sender As Object, e As EventArgs) Handles LikeClear.Click
        LikeBox.Clear()
    End Sub

    Private Sub DontLikeSubmit_Click(sender As Object, e As EventArgs) Handles DontLikeClear.Click
        DontLikeBox.Clear()
    End Sub

    Private Sub FaveSelect_Click(sender As Object, e As EventArgs) Handles Fave.SelectedIndexChanged
        TellFavorite(Fave.Text)
    End Sub

    Private Sub FaveClear_Click(sender As Object, e As EventArgs) Handles FaveCont.Click
        Dim message1 As String = Fave.Text
        Dim message2 As String
        If message1 = "animal" Then
            message2 = "I love lions' roars!"
        ElseIf message1 = "book" Then
            message2 = "I love chocolate!"
        ElseIf message1 = "color" Then
            message2 = "My favorite character in Super Mario Bros is Luigi and he has a green cap"
        ElseIf message1 = "flavor" Then
            message2 = "I could eat chocolate for any meal!"
        ElseIf message1 = "food" Then
            message2 = "Chocolate cake's the best!"
        ElseIf message1 = "game" Then
            message2 = "I just love Luigi"
        ElseIf message1 = "movie" Then
            message2 = "It's about robots, like me!"
        ElseIf message1 = "sport" Then
            message2 =
                "It's the greatest sport in the world.  But did you know they call it football in other countries?"
        Else
            message2 = "I never want summer to end!"
        End If
        SpeakString(message2)
    End Sub

#End Region

#Region "Conversation Repair"

    Private Sub Zoo1_Click(sender As Object, e As EventArgs) Handles Zoo1.Click
        DisableTimers()
        SendCommand("R")
        SpeakString("The zoo is fun.")
        Zoo1.BackColor = Color.Gold
        EnableTimers()
    End Sub

    Private Sub Zoo2_Click(sender As Object, e As EventArgs) Handles Zoo2.Click
        SpeakString("I feel bad for the polar bears.  I hope they don't melt in the zoo, but they're my favorite part!", -2, 100, True)
    End Sub

    Private Sub ImSleepy1_Click(sender As Object, e As EventArgs) Handles ImSleepy1.Click
        SpeakString("I'm sleepy today", 0.5, 100, True)
        ImSleepy1.BackColor = Color.Gold
    End Sub

    Private Sub ImSleepy2_Click(sender As Object, e As EventArgs) Handles ImSleepy2.Click
        SpeakString("Sometimes I get so tired during the day, I just fall asleep", -2, 100, True)
    End Sub

    Private Sub Cookies1_Click(sender As Object, e As EventArgs) Handles Cookies1.Click
        SpeakString("Do you like cookies?  I love cookies", -2, 100, True)
        Cookies1.BackColor = Color.Gold
    End Sub

    Private Sub Cookies2_Click(sender As Object, e As EventArgs) Handles Cookies2.Click
        SpeakString("My favorite kinds are chocolate chip, but I also really like peanut butter cookies.", -2, 100, True)
    End Sub

    Private Sub Friend1_Click(sender As Object, e As EventArgs) Handles Friend1.Click
        SpeakString("Do you know, I've got a friend named Jimmy, and he's really funny.", -2, 100, True)
        Friend1.BackColor = Color.Gold
    End Sub

    Private Sub Friend2_Click(sender As Object, e As EventArgs) Handles Friend2.Click
        SpeakString("One time he made such a silly face, I couldn't stop laughing for five whole minutes!", -2, 100, True)
    End Sub

#End Region

#Region "Weekend"
    Private Sub WkndBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WkndBox.SelectedIndexChanged
        Dim message1, message2 As String
        message1 = WkndBox.Text
        If message1 = "beach" Then
            message2 = "This weekend I went to the beach"
        ElseIf message1 = "birthday party" Then
            message2 = "This weekend was my friend Charlie's birthday.  He had an awesome birthday party"
        ElseIf message1 = "lego" Then
            message2 = "This weekend I played with legos"
        Else
            message2 = "This weekend I played some video games."
        End If
        SpeakString(message2, -2, 100, True)
    End Sub
    Private Sub WkndCont1_Click(sender As Object, e As EventArgs) Handles WkndCont1.Click
        Dim message1 As String = WkndBox.Text
        Dim message2 As String
        If message1 = "beach" Then
            message2 = "I always make sand castles when I go to the beach.  I'm a pro at a making sand castles!"
        ElseIf message1 = "birthday party" Then
            message2 = "We played some fun games and there was a cool robot dance"
        ElseIf message1 = "lego" Then
            message2 = "I built a really cool tower.  Next time I'm gonna build one so tall, it's going to be taller than this table!"
        Else
            message2 = "I played Mario Cart and beat my high score and made it to the next level!"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub
    Private Sub WkndCont2_Click(sender As Object, e As EventArgs) Handles WkndCont2.Click
        Dim message1 As String = WkndBox.Text
        Dim message2 As String
        If message1 = "beach" Then
            message2 = "I collect seashells so I love finding really cool seashells when I'm at the beach"
        ElseIf message1 = "birthday party" Then
            message2 = "I ate really good cake at the party too!"
        ElseIf message1 = "lego" Then
            message2 = "Legos are really fun because you can build anything you want!"
        Else
            message2 = "I love playing racing games. I wish I could watch a car race in real life!"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub
#End Region

#Region "Responses"

    Private Sub DontKnow_Click(sender As Object, e As EventArgs) Handles DontKnow.Click
        SpeakString("I don't know", -2, 100, True)
    End Sub

    Private Sub Yes_Click(sender As Object, e As EventArgs) Handles Yes.Click
        SpeakString("Yes", -2, 100, True)
    End Sub

    Private Sub No_Click(sender As Object, e As EventArgs) Handles No.Click
        SpeakString("No", -2, 100, True)
    End Sub

    Private Sub Sorry_Click(sender As Object, e As EventArgs) Handles Sorry.Click
        SpeakString("Sorry", -2, 100, True)
    End Sub

    Private Sub HBU_Click(sender As Object, e As EventArgs) Handles HBU.Click
        SpeakString("How about you?", -2, 100, True)
    End Sub

    Private Sub Guess_Click(sender As Object, e As EventArgs) Handles MeToo.Click
        SpeakString("Me too", -2, 100, True)
    End Sub

    Private Sub Maybe_Click(sender As Object, e As EventArgs) Handles Maybe.Click
        SpeakString("Maybe", -2, 100, True)
    End Sub

#End Region

#Region "Exclamations"

    Private Sub Exclaim1_Click(sender As Object, e As EventArgs) Handles Exclaim1.Click
        SpeakString("How cool!", -2, 100, True)
    End Sub

    Private Sub Exclaim2_Click(sender As Object, e As EventArgs) Handles Exclaim2.Click
        SpeakString("Wow!", -2, 100, True)
    End Sub

    Private Sub Exclaim3_Click(sender As Object, e As EventArgs) Handles Exclaim3.Click
        SpeakString("That's great!", -2, 100, True)
    End Sub

    Private Sub Exclaim4_Click(sender As Object, e As EventArgs) Handles Exclaim4.Click
        SpeakString("You're AWESOME! " + myName, -2, 100, True)
    End Sub

    Private Sub YoureCool_Click(sender As Object, e As EventArgs) Handles YoureCool.Click
        SpeakString("You're real cool!", -2, 100, True)
    End Sub

#End Region

#Region "Speech Controls"

    Private Sub StopButton_Click(sender As Object, e As EventArgs) Handles StopButton.Click
        Try
            If currentThread.IsAlive Then
                Try
                    currentThread.Abort()
                Catch ex As ThreadStateException
                    currentThread.Resume()
                End Try
            End If
        Catch ex As Exception

        End Try
        stop_Clicked = True
        Select Case speaker.State
            Case SynthesizerState.Speaking
                speaker.SpeakAsyncCancelAll()
                Exit Select
            Case SynthesizerState.Paused
                speaker.Resume()
                speaker.SpeakAsyncCancelAll()
                Exit Select
        End Select
        speaker.SpeakAsyncCancelAll()
        SendCommand("AA")
        DisableTimers()
    End Sub

    Private Sub Pause_Click(sender As Object, e As EventArgs) Handles Pause.Click
        Try
            If currentThread.IsAlive Then
                currentThread.Suspend()
            End If
        Catch ex As Exception
            Exit Try
        End Try
        DisableTimers()
        Select Case speaker.State
            Case SynthesizerState.Speaking
                speaker.Pause()
                Exit Select
        End Select
    End Sub

    Private Sub ResumeButton_Click(sender As Object, e As EventArgs) Handles ResumeButton.Click
        Try
            If currentThread.IsAlive Then
                Try
                    currentThread.Resume()
                Catch ex As ThreadStateException
                    Exit Try
                End Try
            End If
        Catch ex As NullReferenceException
            Exit Try
        End Try
        Select Case speaker.State
            Case SynthesizerState.Paused
                speaker.Resume()
                Exit Select
        End Select
        EnableTimers()
    End Sub

#End Region

#Region "Intro / Exit"

    Private Sub Hello_Click(sender As Object, e As EventArgs) Handles Hello.Click
        SpeakString("Hello.", -2, 100, True)
    End Sub

    Private Sub Hi_Click(sender As Object, e As EventArgs) Handles Hi.Click
        SpeakString("Hi, my name is L-E. What's your name?", -2, 100, True)
    End Sub

    Private Sub MoveDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MoveDrop.SelectedIndexChanged
        DisableTimers()
        Dim selection As String
        selection = MoveDrop.Text
        If selection = "eyes" Then
            SpeakString("Look what I can do with my eyes!!")
            SendCommand("S")
        ElseIf selection = "mouth" Then
            SpeakString("Look what I can do with my mouth!!")
            SendCommand("Z")
        Else
            SpeakString("Look what faces I can make!!")
            SendCommand("Y")
        End If
        EnableTimers()
    End Sub

    Private Sub Bye_Click(sender As Object, e As EventArgs) Handles Bye.Click
        SpeakString("Bye " + myName + ". That was fun!", -2, 100, True)
    End Sub

    Private Sub SeeYou_Click(sender As Object, e As EventArgs) Handles SeeYou.Click
        SpeakString("See you next time!", -2, 100, True)
    End Sub

#End Region

#Region "Story"

    Private Sub Story1_Click_1(sender As Object, e As EventArgs) Handles Story1.Click
        SpeakString("Do you wanna hear a story?", -2, 100, True)
    End Sub

    Private Sub Story2_Click(sender As Object, e As EventArgs) Handles Story2.Click
        SpeakString(
            "One time, I was at a pool party.  It was my friend Jimmy's birthday, and he decided we should play a game. " &
            "He dropped these jewels at the bottom of the pool, and you had to see how many you could collect in thirty seconds. " &
            "Most people got around five.  But my friend Andy?  He was underwater for so long, we weren't sure if something had happened! " &
            "We were about to send someone in to look for him, when all of a sudden, he popped out of the pool.  He said he saw something " &
            "sparkling at the bottom and thought it was a jewel, but it was covered under some dirt at the bottom of the pool, so he had to " &
            "scrape it out.  And guess what?  He found a necklace Jimmy's sister dropped in the pool a long time ago!  We were all looking for " &
            "Jewels, but, really, it was Andy who found the real treasure!", -3)
    End Sub

    Private Sub Story3_Click(sender As Object, e As EventArgs) Handles Story3.Click
        SpeakString("Do you want to hear another story?", -3, 100, True)
    End Sub

    Private Sub Story4_Click(sender As Object, e As EventArgs) Handles Story4.Click
        SpeakString(
            "Once, I had a playdate at my friend Alex's house and we really wanted to bake something. We both liked cookies and pizza, " &
            "so we thought it would be an awesome idea to bake a cookie pizza! I used chocolate chips, mini oreos, marshmallows, and peanut " &
            "butter chips as my toppings. Alex made a crazy pizza with gummy worms, M and M's, and bacon. Those were the weirdest cookie toppings " &
            "I had ever seen. We put our cookie pizzas in the oven, and once they were done, we were so hungry, we ate 4 slices! I tried some of Alex's " &
            "cookie pizza, and it was actually pretty good. Maybe I'll try making a crazy cookie pizza next time. That was one of the best playdates " &
            "I've ever had.")
    End Sub

    Private Sub Story5_Click(sender As Object, e As EventArgs) Handles Story5.Click
        SpeakString("Why don't you tell me a story!", -2, 100, True)
    End Sub

#End Region

#Region "General Functions"

    Private Sub EnableTimers()
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
    End Sub

    Private Sub DisableTimers()
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
    End Sub

    Private Sub SendCommand(cmd As String, Optional resetTimers As Boolean = False)
        If failedToConnect Then Exit Sub
        If resetTimers Then DisableTimers()
        SerialPort1.Open()
        SerialPort1.Write(cmd)
        SerialPort1.Close()
        If resetTimers Then EnableTimers()
    End Sub

    Private Sub SpeakString(txt As String, Optional rate As Double = -2, Optional volume As Double = 100,
                            Optional resetTimers As Boolean = False)
        If ResetTimers Then DisableTimers()
        speaker.Rate = rate
        speaker.Volume = volume
        speaker.SpeakAsync(txt)
        If ResetTimers Then EnableTimers()
    End Sub

    Private Sub TellFavorite(type As String)
        Dim message2 As String
        If type = "animal" Then
            message2 = "a lion"
        ElseIf type = "book" Then
            message2 = "Charlie and the Chocolate Factory"
        ElseIf type = "color" Then
            message2 = "green"
        ElseIf type = "flavor" Then
            message2 = "chocolate"
        ElseIf type = "food" Then
            message2 = "cake"
        ElseIf type = "game" Then
            message2 = "Super Mario Bros."
        ElseIf type = "movie" Then
            message2 = "Wall-E"
        ElseIf type = "sport" Then
            message2 = "soccer"
        Else
            message2 = "Phineas and Ferb"
        End If
        SpeakString("My favorite " + type + "is " + message2, -2, 100, True)
    End Sub

#End Region

#Region "Speech Recognition"
    'I tried to group all the speech recognition code together but there were some places that I couldn't.
    'Just look for '#JC
    Private ReadOnly DefaultResponses As New List(Of String)
    Private KnownResponses As Dictionary(Of String, String)

    Private ResponseHandler As EventHandler
    Private RecoginizedPrompt As String = ""

    Private Speech As String

    'Quick Recognition Variables
    Private PromptHandled As Boolean = False
    Private LastSpeechUpdate As Long = Environment.TickCount
    Private CurrentPrompt As String = ""
    Private HandledSpeech As String = ""
    Private ReadOnly ActionsTaken As List(Of String) = New List(Of String)
    Private NewText As String = ""
    Private ReadOnly SpeechLock As New Object

    'Network Socket for Chrome communication
    ReadOnly allSockets As New List(Of IWebSocketConnection)()
    ReadOnly server As New WebSocketServer("wss://0.0.0.0:8182")

    Sub InitSpeechRecoginition()

        'Setup the websocket server to receive connections from Chrome
        server.Certificate = New X509Certificate2("fritz.pfx", "oursslpass")

        server.Start(Function(socket)
                         socket.OnOpen = Function()
                                             allSockets.Add(socket)
                                         End Function
                         socket.OnClose = Function()
                                              allSockets.Remove(socket)
                                          End Function
                         socket.OnMessage = Function(message)
                                                If message <> "Connected Successfully!" Then
                                                    SpeechRecognized(message)
                                                Else
                                                    Speech = "(Re)Connected to Chrome dictation page successfully."
                                                End If
                                            End Function
                     End Function)


        'Load the AutomationActions Before Starting to Listen
        'These map potential phrases to actions. They could be improved however
        'in order to follow a script, or not repeat questions, etc.

        'LOOK BELOW THIS SUB TO MAP ACTIONS TO THESE RESPONSES

        'Introductions
        DefaultResponses.Add("hello l-e")
        DefaultResponses.Add("hi l-e")
        DefaultResponses.Add("hi")
        DefaultResponses.Add("hello")

        'Small Talk
        DefaultResponses.Add("how are you")
        DefaultResponses.Add("how're you doing")
        DefaultResponses.Add("what did you do this weekend")
        DefaultResponses.Add("what did you do today")

        'Stories
        DefaultResponses.Add("tell me a story")

        'Favorites
        DefaultResponses.Add("what is your favorite meal")
        DefaultResponses.Add("what is your favorite animal")
        DefaultResponses.Add("what is your favorite book")
        DefaultResponses.Add("what is your favorite color")
        DefaultResponses.Add("what is your favorite flavor")
        DefaultResponses.Add("what is your favorite food")
        DefaultResponses.Add("what is your favorite game")
        DefaultResponses.Add("what is your favorite movie")
        DefaultResponses.Add("what is your favorite tv show")

        'Special Cases
        DefaultResponses.Add("yes")
        DefaultResponses.Add("no")
        DefaultResponses.Add("why")

        'LOOK BELOW THIS SUB TO MAP ACTIONS TO THESE RESPONSES

        ResponseHandler = AddressOf HandleResponse

        KnownResponses = LoadDictionaryFromCsv("responses.csv")
        If (KnownResponses Is Nothing) Then KnownResponses = New Dictionary(Of String, String)
    End Sub

    Private Sub HandleResponse(sender As Object, e As EventArgs)
        Select Case (RecoginizedPrompt)

            'Introductions
            Case "hello"
                Hi_Click(Nothing, Nothing)
            Case "hello l-e"
                Hello_Click(Nothing, Nothing)
            Case "hi"
                Hi_Click(Nothing, Nothing)
            Case "hi l-e"
                Hello_Click(Nothing, Nothing)

                'Small Talk
            Case "how are you"
                NotGreat_Click(Nothing, Nothing)
            Case "how are you doing"
                NotGreat_Click(Nothing, Nothing)
            Case "what did you do this weekend"
                'Lego1_Click(Nothing, Nothing)
            Case "what did you do today"
                TodayA_Click(Nothing, Nothing)

                'Stories
            Case "tell me a story"
                Story2_Click(Nothing, Nothing)

                'Favorites
            Case "what is your favorite meal"
                TellFavorite("food")
            Case "what is your favorite animal"
                TellFavorite("animal")
            Case "what is your favorite book"
                TellFavorite("book")
            Case "what is your favorite color"
                TellFavorite("color")
            Case "what is your favorite flavor"
                TellFavorite("flavor")
            Case "what is your favorite food"
                TellFavorite("food")
            Case "what is your favorite game"
                TellFavorite("game")
            Case "what is your favorite movie"
                TellFavorite("movie")
            Case "what is your favorite tv show"
                TellFavorite("tv show")

                'Special Cases
            Case "why"
                NotGreatCont_Click(Nothing, Nothing)
        End Select
    End Sub

    'This is called automatically whenever speech is successfully recognized.
    Public Sub SpeechRecognized(heardtext As String)
        Dim matchThreshhold = 0.65
        Dim currentMatchVal = 0.0
        Dim currentMatch As String
        Dim finalized As Boolean
        Dim time As Long = Environment.TickCount

        'This is a delegate for any action we may take.
        If heardtext = "" Then Exit Sub
        SyncLock SpeechLock
            'Because we are in a different thread, we update a class variable which a timer sets the form to display (We should use a delegate here)
            If (HandledSpeech.Length > 0 And heardtext.Length > 0) Then heardtext = heardtext.Replace(HandledSpeech, "")
            If heardtext.Length > 2 Then
                If (NewText <> heardtext) Then LastSpeechUpdate = time
                If heardtext.Substring(0, 2) = "%1" Then
                    Speech = heardtext.Substring(2).Trim()
                    Speech = Speech.Replace("  ", " ").Trim()
                    finalized = True
                    PromptHandled = False
                ElseIf heardtext.Substring(0, 2) = "%0" Then
                    Speech = heardtext.Substring(2).Trim()
                    Speech = Speech.Replace("  ", " ").Trim()
                    PromptHandled = False
                    NewText = Speech
                End If
            End If
            Speech = Speech.Replace("  ", " ")
            If ((CurrentPrompt = Speech And LastSpeechUpdate + 500 < time) Or (finalized)) And PromptHandled = False _
                Then
                'Check for common misconceptions
                Speech = Speech.ToLower()
                Speech = Speech.Replace("ellie", "l-e")
                Speech = Speech.Replace("allie", "l-e")
                Speech = Speech.Replace(" y ", " why ") 'Another common misconception that I am course correcting.
                Speech = Speech.Trim()

                'If we are accepting automatic actions, score the phrase against all our potential responses, find the best and implement the action.
                If chkAutoAction.Checked Then
                    If (GetIndexOfKey(KnownResponses, Speech) > -1) Then
                        Speech = KnownResponses(Speech)
                    End If
                    For Each p In DefaultResponses
                        If (Compare(p, Speech) > currentMatchVal) Then
                            currentMatchVal = Compare(p, Speech)
                            currentMatch = p
                        End If
                    Next
                    If (currentMatchVal >= matchThreshhold) Then
                        If (speaker.State = SynthesizerState.Ready And ActionsTaken.IndexOf(currentMatch) = -1) Then
                            RecoginizedPrompt = currentMatch
                            ResponseHandler.Invoke(Nothing, Nothing)
                            ActionsTaken.Add(currentMatch)
                        End If
                        If (GetIndexOfKey(KnownResponses, Speech) = -1) Then KnownResponses.Add(Speech, currentMatch)
                    Else
                        If (GetIndexOfKey(KnownResponses, Speech) = -1) Then KnownResponses.Add(Speech, "")
                    End If
                    SaveDictionaryToCsv("responses.csv", KnownResponses)
                End If
                PromptHandled = True
                If Not finalized Then
                    HandledSpeech = heardtext
                Else
                    HandledSpeech = ""
                End If
                If finalized Then ActionsTaken.Clear()
            Else
                CurrentPrompt = Speech
            End If
        End SyncLock
    End Sub

    Function Compare(str1 As String, str2 As String) As Double
        Compare = Distance(str1, str2)
    End Function

    Private Function Distance(s1 As String, s2 As String) As Single
        Dim p1 = GetPairs(s1)
        Dim p2 = GetPairs(s2)
        Return (2.0F * p1.Intersect(p2).Count()) / (p1.Count + p2.Count)
    End Function

    Private Function GetPairs(s As String) As List(Of String)
        If s Is Nothing Then
            Return New List(Of String)()
        End If
        If s.Length < 3 Then
            Return New List(Of String)() From { _
                s _
                }
        End If

        Dim result As New List(Of String)()
        For i = 0 To s.Length - 2
            result.Add(s.Substring(i, 2).ToLower(CultureInfo.InvariantCulture))
        Next
        Return result
    End Function
    'This timer updates our form with the latest speech if any.
    Private Sub tmrSpeech_Tick(sender As Object, e As EventArgs) Handles tmrSpeech.Tick
        lblSpeechHeard.Text = Speech
        If PromptHandled Then
            lblSpeechHeard.ForeColor = Color.Blue
        Else
            lblSpeechHeard.ForeColor = Color.Red
        End If
        SpeechRecognized(CurrentPrompt)
    End Sub

    Private Sub SaveDictionaryToCsv(Filename As String, Dict As Dictionary(Of String, String))
        Dim outFile As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(Filename, False)
        For i = 0 To Dict.Keys.Count - 1
            outFile.WriteLine(Dict.Keys(i) & "," & Dict(Dict.Keys(i)))
        Next
        outFile.Close()
    End Sub

    Private Function LoadDictionaryFromCsv(Filename As String) As Dictionary(Of String, String)
        Dim returnDict As New Dictionary(Of String, String)
        Try
            If File.Exists(Filename) Then
                Using MyReader As New TextFieldParser(Filename)
                    MyReader.Delimiters = {","}
                    Dim currentRow As String()
                    While Not MyReader.EndOfData
                        Try
                            currentRow = MyReader.ReadFields()
                            If (currentRow.Count = 2) Then
                                returnDict.Add(currentRow(0), currentRow(1))
                            End If
                        Catch ex As MalformedLineException
                            MsgBox("Line " & ex.Message &
                                   "is not valid and will be skipped.")
                        End Try
                    End While
                End Using
            End If
        Catch ex As Exception
        End Try
        LoadDictionaryFromCsv = returnDict
    End Function

    Private Function GetIndexOfKey(tempDict As Dictionary(Of String, String), key As [String]) As Integer
        Dim index As Integer = -1
        For Each value As [String] In tempDict.Keys
            index += 1
            If key = value Then
                Return index
            End If
        Next
        Return -1
    End Function

#End Region

#Region "School"
    Private Sub SchoolBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SchoolBox.SelectedIndexChanged
        Dim message1, message2 As String
        message1 = SchoolBox.Text
        If message1 = "fave subject" Then
            message2 = "My favorite subject is English"
        ElseIf message1 = "least fave subject" Then
            message2 = "My least favorite subject is probably history"
        ElseIf message1 = "recess" Then
            message2 = "I love recess"
        ElseIf message1 = "teacher" Then
            message2 = "My teacher's pretty nice"
        ElseIf message1 = "school's okay" Then
            message2 = "School's okay"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub
    Private Sub SchoolCont1_Click(sender As Object, e As EventArgs) Handles SchoolCont1.Click
        Dim message1, message2 As String
        message1 = SchoolBox.Text
        If message1 = "fave subject" Then
            message2 = "I really love reading books"
        ElseIf message1 = "least fave subject" Then
            message2 = "I find history really boring"
        ElseIf message1 = "recess" Then
            message2 = "Playing in the playground is always fun!"
        ElseIf message1 = "teacher" Then
            message2 = "She sometimes bakes for us"
        ElseIf message1 = "school's okay" Then
            message2 = "Sometimes it can be a little boring, but mostly it's not bad"
        End If
        If String.IsNullOrEmpty(message2) Then
            SpeakString(message2, -2, 100, True)
        End If
    End Sub

    Private Sub SchoolCont2_Click(sender As Object, e As EventArgs) Handles SchoolCont2.Click
        Dim message1, message2 As String
        message1 = SchoolBox.Text
        If message1 = "fave subject" Then
            message2 = "My favorite is when my teacher reads us adventure stories!"
        ElseIf message1 = "least fave subject" Then
            message2 = "It's not fun memorizing facts about old people"
        ElseIf message1 = "recess" Then
            message2 = "I think the monkey bars are really cool"
        ElseIf message1 = "teacher" Then
            message2 = "Sometimes, if we're really good, she also lets us have a longer recess and snacktime!"
        ElseIf message1 = "school's okay" Then
            message2 = "Homework is the most annoying part, but learning is fun. I want to be really smart one day!"
        End If
        If String.IsNullOrEmpty(message2) Then
            SpeakString(message2, -2, 100, True)
        End If
    End Sub
#End Region

#Region "Summer"
    Private Sub SummerBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SummerBox.SelectedIndexChanged
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Dim message1, message2 As String
        message1 = SummerBox.Text
        If message1 = "pool" Then
            message2 = "In the summer, I love going to the pool"
        ElseIf message1 = "ice cream" Then
            message2 = "My favorite part about summer is eating a lot of ice cream"
        ElseIf message1 = "vacation" Then
            message2 = "I love going on vacation over the summer"
        Else
            message2 = "I love playing games and sports in the summer"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub

    Private Sub SummerCont1_Click(sender As Object, e As EventArgs) Handles SummerCont1.Click
        Dim message1, message2 As String
        message1 = SummerBox.Text
        If message1 = "pool" Then
            message2 = "I can't really swim, but I love sitting on a pool float and floating around in the water"
        ElseIf message1 = "ice cream" Then
            message2 = "I really love any kind of chocolate flavors.  Plain chocolate is probably my favorite"
        ElseIf message1 = "vacation" Then
            message2 = "One day, I really want to go to Hawaii!"
        Else
            message2 = "I love soccer.  It's a really cool sport. And fun to watch!"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub

    Private Sub SummerCont2_Click(sender As Object, e As EventArgs) Handles SummerCont2.Click
        Dim message1, message2 As String
        message1 = SummerBox.Text
        If message1 = "pool" Then
            message2 = "You always have to be careful, though, and make sure there are no thunderstorms before you decide to go in the pool!"
        ElseIf message1 = "ice cream" Then
            message2 = "Oh, and I love ice cream in cones with sprinkles on top.  Sprinkles make everything better!"
        ElseIf message1 = "vacation" Then
            message2 = "I'm kind of scared of planes, but travelling is really fun"
        Else
            message2 = "It's really nice not having homework in the summer and just being able to play!"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub
#End Region

#Region "Favorites"
    Private Sub FaveBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FaveBox.SelectedIndexChanged
        Dim message1 As String
        message1 = FaveBox.Text
        SpeakString("What is your favorite" + message1, -2, 100, True)
    End Sub

    Private Sub FaveAns_Click(sender As Object, e As EventArgs) Handles FaveAns.Click
        Dim message1, message2 As String
        message1 = FaveBox.Text
        If message1 = "animal" Then
            message2 = "the lion"
        ElseIf message1 = "book" Then
            message2 = "Charlie and the Chocolate Factory"
        ElseIf message1 = "color" Then
            message2 = "green"
        ElseIf message1 = "flavor" Then
            message2 = "chocolate"
        ElseIf message1 = "food" Then
            message2 = "cake"
        ElseIf message1 = "game" Then
            message2 = "Super Mario Bros."
        ElseIf message1 = "movie" Then
            message2 = "Wall-E"
        ElseIf message1 = "sport" Then
            message2 = "soccer"
        ElseIf message1 = "TV show" Then
            message2 = "Phineas and Ferb"
        ElseIf message1 = "part of school" Then
            message2 = "recess"
        ElseIf message1 = "part of summer" Then
            message2 = "eating a lot of ice cream"
        End If
        If Not String.IsNullOrEmpty(message2) Then
            SpeakString("My favorite " + message1 + " is " + message2, -2, 100, True)
        End If
    End Sub

    Private Sub FaveAnsCont_Click(sender As Object, e As EventArgs) Handles FaveAnsCont.Click
        Dim message1 As String = FaveBox.Text
        Dim message2 As String
        If message1 = "animal" Then
            message2 = "I love lions' roars!"
        ElseIf message1 = "book" Then
            message2 = "I love chocolate!"
        ElseIf message1 = "color" Then
            message2 = "My favorite character in Super Mario Bros is Luigi and he has a green cap"
        ElseIf message1 = "flavor" Then
            message2 = "I could eat chocolate for any meal!"
        ElseIf message1 = "food" Then
            message2 = "Chocolate cake's the best!"
        ElseIf message1 = "game" Then
            message2 = "I just love Luigi"
        ElseIf message1 = "movie" Then
            message2 = "It's about robots, like me!"
        ElseIf message1 = "sport" Then
            message2 = "It's the greatest sport in the world.  But did you know they call it football in other countries?"
        ElseIf message1 = "TV show" Then
            message2 = "I never want summer to end!"
        ElseIf message1 = "part of school" Then
            message2 = "It's the most relaxing part of the day since it's so fun!"
        ElseIf message1 = "part of summer" Then
            message2 = "Ice cream is the greatest on a hot summer day!"
        End If
        If Not String.IsNullOrEmpty(message2) Then
            SpeakString(message2, -2, 100, True)
        End If
    End Sub
#End Region

#Region "Conversation"
    Private Sub ConvoBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ConvoBox.SelectedIndexChanged
        Dim message1, message2 As String
        message1 = ConvoBox.Text
        If message1 = "Like to do at home?" Then
            message2 = "What do you like to do at home?"
        ElseIf message1 = "Best gift ever given to you?" Then
            message2 = "What was the best gift you ever received?"
        ElseIf message1 = "Favorite place in world?" Then
            message2 = "What is your favorite place in the world?"
        ElseIf message1 = "If could invent something, what?" Then
            message2 = "If you could invent something, what would it be?"
        ElseIf message1 = "Funniest thing that happened?" Then
            message2 = "What is the funniest thing that ever happened to you?"
        ElseIf message1 = "Imagine the future..." Then
            message2 = "What do you think the future will be like?"
        ElseIf message1 = "If could be animal, which one?" Then
            message2 = "If you could be an animal, which one would you be?"
        ElseIf message1 = "What want to be when grow up?" Then
            message2 = "If you could be anything you want to be when you grow up, what would it be?"
        End If
        SpeakString(message2, -2, 100, True)
    End Sub

    Private Sub ConvoAns_Click(sender As Object, e As EventArgs) Handles ConvoAns.Click
        Dim message1, message2 As String
        message1 = ConvoBox.Text
        If message1 = "Like to do at home?" Then
            message2 = "I like playing video games and reading"
        ElseIf message1 = "Best gift ever given to you?" Then
            message2 = "The best gift I ever received was a robot, like me, called Room-ba.  It vacuums your room for you so you don't have to!"
        ElseIf message1 = "Favorite place in world?" Then
            message2 = "My favorite place is New Zealand.  It's the most beautiful place in the world!"
        ElseIf message1 = "If could invent something, what?" Then
            message2 = "I would invent a machine that would give you whatever food you want at the click of a button! I just love food"
        ElseIf message1 = "Funniest thing that happened?" Then
            message2 = "Well it didn't really happen to me, but I have a friend who's from a different country, so when she first moved to America, she didn't really understand English too well.  And our teacher wanted us to write about the Superbowl.  And she didn't know that the Superbowl was a football game since she's not from America.  She thought she was supposed to write about a bowl that was super.  So she created an entire story about a bowl with magical powers.  And she was the first one to volunteer to read her paper the next day, so we were all cracking up when she read it out loud because she had no idea that the Superbowl is about football!"
        ElseIf message1 = "Imagine the future..." Then
            message2 = "I think the future will be really cool and there will be lots of robots like me! And the robots will do all sorts of stuff for the humans, so they don't have to!"
        ElseIf message1 = "If could be animal, which one?" Then
            message2 = "If I could be an animal, I'd be a bird because I think it would be so cool to fly!"
        ElseIf message1 = "What want to be when grow up?" Then
            message2 = "When I grow up, I think I want to be a doctor!  It's pretty cool to cure sick people!"
        End If
        If Not String.IsNullOrEmpty(message2) Then
            SpeakString(message2, -2, 100, True)
        End If
    End Sub
    Private Sub TellMore_Click(sender As Object, e As EventArgs) Handles TellMore.Click
        SpeakString("Tell me more!", -2, 100, True)
    End Sub

    Private Sub Why_Click(sender As Object, e As EventArgs) Handles Why.Click
        SpeakString("Why", -2, 100, True)
    End Sub

#End Region
    'Private Sub Yawn_Click(sender As Object, e As EventArgs) Handles Yawn.Click
    '    Timer1.Enabled = False
    '    SerialPort1.Open()
    '    SerialPort1.Write("O")
    '    SerialPort1.Close()
    '    My.Computer.Audio.Play(My.Resources.Elves, AudioPlayMode.WaitToComplete)

    'End Sub

    'Private Sub Burp_Click(sender As Object, e As EventArgs) Handles Burp.Click
    '    Timer1.Enabled = False
    '    SerialPort1.Open()
    '    SerialPort1.Write("P")
    '    SerialPort1.Close()
    '    My.Computer.Audio.Play(My.Resources.Elves, AudioPlayMode.WaitToComplete)
    'End Sub
End Class
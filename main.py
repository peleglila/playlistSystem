# encoding=utf8
from datetime import datetime, timedelta
import time
import json
import urllib2
import xlrd
from xlutils.copy import copy

# Declare the files path
ratingDataFile = "F:/tests/YES/test.xls"
#ratingDataFile = "M:/yesLineup/test.xls"
playlistFile = "F:/tests/YES/playlist.xls"

# Error mail details:
smtpHost = 'smtp-pulse.com'
smtpPort = 2525
smtpUser = 'pelegl@promots.tv'
smtpPassword = 'rE32iXgmFfjX'
smtpSendList = ['pelegalila@gmail.com', 'pelegl@promots.tv']
# timeToWaitMail in SECONDS
timeToWaitMail = 0.2

# Xlrd workbook declaration
ratingBook = xlrd.open_workbook(ratingDataFile)
playlistBook = xlrd.open_workbook(playlistFile)
playlistCopy = copy(playlistBook)

# Playlist Xls variable
titleColumn = 18
ratingColumn = 31
recommendationColumn = 32

# Gets the tile_ID's from the playlist file
playlistSheet = playlistBook.sheets()[0]
playlistTitleIDs = [int(playlistSheet.cell_value(i, 18)) for i in range(1, playlistSheet.nrows)]

# Take the recommendation information
ratingSheet = ratingBook.sheets()[0]

orcaConnection = True
ratingDict = {}
newRatingStructure = {}

top1day = None
top1week = None
top1month = None
top5week = list()
top10week = list()
top10month = list()

# Definition of the Day, Week and month date to take from
day_ago = datetime.now() - timedelta(days=15)
week_ago = datetime.now() - timedelta(days=22)
month_ago = datetime.now() - timedelta(days=46)

# Channels list
channels = {'yes1 HD': {'playlistChannel': 'YES1', 'displayChannel': 'yes1'},
            'yes2 HD': {'playlistChannel': 'YES2', 'displayChannel': 'yes2'},
            'yes3 HD': {'playlistChannel': 'YES3', 'displayChannel': 'yes3'},
            'yes4 HD': {'playlistChannel': 'YES4', 'displayChannel': 'yes4'},
            'yes5 HD': {'playlistChannel': 'YES5', 'displayChannel': 'yes5'}}


# Colors of STOUT
class Colors:
    def __init__(self):
        pass

    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def sendErrorMail():
    from smtplib import SMTP
    import datetime
    import socket
    localIp = socket.gethostbyname(socket.gethostname())
    debuglevel = 0

    smtp = SMTP()
    smtp.set_debuglevel(debuglevel)
    smtp.connect(smtpHost, smtpPort)
    smtp.login(smtpUser, smtpPassword)

    from_addr = "YES Playlist System <pelegl@promots.tv>"
    to_addr = smtpSendList

    subj = "ERROR : Playlist file is open - SCRIPT CAN\'T RUN"
    date = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    message_text = "-------------------------- ERROR -------------------------\n\n" \
                   "date: %s\n"\
                   "This is a mail from your YES playlist system.\n\n" \
                   "On IP: %s\n\n"\
                   "The file location:\n%s\n\n" \
                   "Is open and the Rating script can not run!\n" \
                   "Please close it and RUN THE SCRIPT AGAIN.\n\n" \
                   "-------------------------- ERROR -------------------------\n\n" \
                   "Thank you,\nPromotheus" % (date, localIp, playlistFile)

    msg = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n%s" % (from_addr, to_addr, subj, date, message_text)

    smtp.sendmail(from_addr, to_addr, msg)
    smtp.quit()


# Main function
def main():
    ratingStructure()
    # for channel in channels:
    #    print channel + ': Day - %s' %  (topChannelRating(channel, day_ago))
    #    print channel + ': Week - %s' % (topChannelRating(channel, week_ago))
    #    print channel + ': Month - %s' % (topChannelRating(channel, month_ago))
    topThisMonth()
    topThisWeek()
    topThisDay()
    checkData(playlistTitleIDs)


# Read Data from O.R.C.A Api
def readXml(titleID):
    global orcaConnection
    url = 'http://217.109.104.6/compass/GetRecommendationListByPopularContributor?external_content_id=' + str(
        titleID) + '&client=json'
    try:
        jsonData = json.load(urllib2.urlopen(url))
    except urllib2.URLError:
        print '\n{0}--- Could not connect to ORCA API, Please contact the Administrator! ---'.format(Colors.FAIL)
        orcaConnection = False
        return False

    recommendedTitle = None
    for data in jsonData['response']:
        if data == 'status' or data == "":
            recommendedTitle = None
            break
        else:
            #            print data['name']
            recommendedTitle = data['name']
    return recommendedTitle


# check if there is any recommendation for the title
def checkRecommendations(titleId, row):
    global orcaConnection
    if orcaConnection:
        print 'what!?'
        recommendedTitle = readXml(titleId)
        writeToExcel(recommendedTitle, row, 32)



# Check in what rating criterion the title is
def checkRating(titleIds, row, titleChannel):
    global rating
    for index in range(0, 1):
        if titleIds == top1day:
            rating = 'הסרט הנצפה ביותר ב24 שעות האחרונות'
            break
        elif titleIds == top1week:
            rating = 'הסרט הכי נצפה השבוע'
            break
        elif titleIds == top1month:
            rating = 'הסרט הכי נצפה החודש'
            break
        elif titleIds == topChannelRating(titleChannel, day_ago):
            rating = 'הסרט הכי נצפה היום בערוץ ' + str(channels[titleChannel]['displayChannel'])
            break
        elif titleIds == topChannelRating(titleChannel, week_ago):
            rating = 'הסרט הכי נצפה השבוע בערוץ ' + str(channels[titleChannel]['displayChannel'])
            break
        elif titleIds == topChannelRating(titleChannel, month_ago):
            rating = 'הסרט הכי נצפה החודש בערוץ ' + str(channels[titleChannel]['displayChannel'])
            break
        elif titleIds in top5week:
            rating = 'אחד מחמשת הסרטים הנצפים השבוע'
            break
        elif titleIds in top10week:
            rating = 'אחד מעשרת הסרטים הנצפים השבוע'
            break
        elif titleIds in top10month:
            rating = 'אחד מעשרת הסרטים הנצפים החודש'
            break
        else:
            rating = None
    writeToExcel(rating, row, ratingColumn)


# Check if title_IDs exists in one of the lists
def checkData(titleIds):
    titleIndex = 1
    for title_id in titleIds:
        checkRecommendations(title_id, titleIndex)
        if title_id in newRatingStructure:
            titleChannel = newRatingStructure[title_id]['channel']
            checkRating(title_id, titleIndex, titleChannel)
            titleIndex += 1
        else:
            print Colors.WARNING + 'Title ID %s - not in the rating list' % title_id
            titleIndex += 1
            continue
    saveExcel()


# Add the rating value to the playlist file
def writeToExcel(rating, row, column):
    copySheet = playlistCopy.get_sheet(0)
    copySheet.write(0, ratingColumn, 'Rating')
    copySheet.write(0, recommendationColumn, 'Recommendation')

    if (rating is not None) and (column == ratingColumn):
        copySheet.write(row, column, unicode(rating, 'utf-8'))
    elif rating is not None:
        copySheet.write(row, column, rating)


# Save playlist file
def saveExcel():
    counter = 1
    while True:
        if counter >= (timeToWaitMail * 60) / 2:
            sendErrorMail()
            break
        try:
            open(playlistFile, "r+")
        except IOError:
            print '\n{0}--- Could not Save the file! Please close Excel! ---'.format(Colors.FAIL)
            counter += 1
            time.sleep(2)
            continue
        break

    playlistCopy.save(playlistFile)
    print '\n{0}--- Writing to Playlist succeed! ---'.format(Colors.OKGREEN)


# Create the channel rating list
def topChannelRating(channel, period):
    periodIndex = periodList(period)
    tempRating = {}
    for title_id in periodIndex:
        if newRatingStructure[title_id]['channel'] == channel:
            rating = newRatingStructure[title_id]['rating']
            tempRating.setdefault(rating, {'title_id': title_id})
        else:
            continue
    tempID = tempRating[max(tempRating)]['title_id']
    return tempID


# Create month top rating lists
def topThisMonth():
    thisMonthListIndex = periodList(month_ago)
    tempRating = {}
    for title_id in thisMonthListIndex:
        rating = newRatingStructure[title_id]['rating']
        tempRating.setdefault(rating, {'title_id': title_id})
    for i in range(0, 10):
        tempID = tempRating[max(tempRating)]['title_id']
        #        tempID = newRatingList.index(str(max(tempRating)))

        if i == 0:
            global top1month
            top1month = tempID
        else:
            top10month.append(tempID)
        tempRating.pop(max(tempRating), None)


# Create week top rating lists
def topThisWeek():
    thisWeekListIndex = periodList(week_ago)
    tempRating = {}
    for title_id in thisWeekListIndex:
        rating = newRatingStructure[title_id]['rating']
        tempRating.setdefault(rating, {'title_id': title_id})
    for i in range(0, 10):
        tempID = tempRating[max(tempRating)]['title_id']
        #        tempID = newRatingList.index(str(max(tempRating)))
        if i == 0:
            global top1week
            top1week = tempID
        elif i <= 4:
            top5week.append(tempID)
        else:
            top10week.append(tempID)
        tempRating.pop(max(tempRating), None)


# Create day top rating title
def topThisDay():
    thisDayListIndex = periodList(day_ago)
    tempRating = {}
    for title_id in thisDayListIndex:
        rating = newRatingStructure[title_id]['rating']
        tempRating.setdefault(rating, {'title_id': title_id})
    for i in range(0, 1):
        tempID = tempRating[max(tempRating)]['title_id']
        if i == 0:
            global top1day
            top1day = tempID


# Return list with indexes of movies from a given date
def periodList(date):
    tempIdList = list()

    for title_id in newRatingStructure:

        tempDataValue = datetime(*xlrd.xldate_as_tuple((newRatingStructure[title_id]['date']), 0))

        if date <= tempDataValue:
            tempIdList.append(title_id)

    return tempIdList


# Create the rating list from the rating file and sum the rating
def ratingStructure():
    tempID = list()
    for i in range(1, ratingSheet.nrows):
        title_id = int(ratingSheet.cell_value(i, 5))
        name = ratingSheet.cell_value(i, 7)
        date = ratingSheet.cell_value(i, 8)
        channel = ratingSheet.cell_value(i, 4)
        rating = ratingSheet.cell_value(i, 11)
        if title_id in tempID:
            newRatingStructure[title_id]['rating'] = newRatingStructure[title_id]['rating'] + rating
        else:
            newRatingStructure.setdefault(title_id,
                                          {'name': name, 'date': date, 'channel': channel, 'rating': rating, 'id': i})
        tempID.append(title_id)


main()

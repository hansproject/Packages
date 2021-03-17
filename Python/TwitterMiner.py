# Pandas is used to create, read, manipulate csv or xlsx files or dataframes
import pandas as pd
import requests
import bs4
import tweepy
import datetime
import xlsxwriter
import sys
from bs4 import BeautifulSoup
from requests_oauthlib import OAuth1
# re library is used to filter data in a tweet or string
import re
import io
import csv
from tweepy import OAuthHandler
from textblob import TextBlob
import numpy as np
# Matplot lib is used to visualize the data
import matplotlib.pyplot as plt

# sklearn for data preprocessing and normalization
from sklearn.preprocessing import MinMaxScaler

# Numpy is used for data manipulation
import numpy as np

# create_exccel_file function which create an Excel file
# Accepts 2 arguments
# @filename:   name of file that will be created
# @raw_tweets: Raw tweets to be written in the file
#------------------------------------------------------------
def createExcelFile(file_name, raw_tweets):
    workbook = xlsxwriter.Workbook(file_name + ".xlsx")
    worksheet = workbook.add_worksheet()
    row = 0

    for tweet in raw_tweets:
        worksheet.write_string(row, 0, str(tweet.id))
        worksheet.write_string(row, 1, str(tweet.created_at))
        worksheet.write(row, 2, tweet.text)
        worksheet.write_string(row, 3, str(tweet.in_reply_to_status_id))
        row += 1

    workbook.close()
    print("Excel file ready")


# tweet_analysis function which return a dictionary that contains tweets polarity per product name
# Accepts 3 arguments
# @polarity:        If True it will return the polarity per product name; if False it will return sentiments per product name
# @product_names:   Name of products
# @health_products: Health products' dictionary
# ------------------------------------------------------------
def tweetsAnalysis(polarity, product_names, health_products):
    if polarity:
        all_prod_tweets_polarity = []
        for name in product_names:
            cur_tweets_polarity = []
            for tweet in health_products[name]:
                # Add 1 polarity value to the list
                cur_tweets_polarity.append(getPolarity(tweet))

                # Add a polarity to tweets_polarity list
                cur_tweets_polarity.append(getPolarity(tweet))
                # Append current tweets sentiments and polarity to the total for each product
                all_prod_tweets_polarity.append(cur_tweets_polarity)
        return {keys: values for keys, values in zip(product_names, all_prod_tweets_polarity)}
    # Else return sentiments if tweets array given
    # --------------------------------------------------------------------
    else:
        all_prod_tweets_sentiments = []
        for name in product_names:
            cur_sentiments = []
            for tweet in health_products[name]:
                # Add a sentiment to sentiments list
                cur_sentiments.append(getTweetSentiment(tweet))

                # Add a polarity to tweets_polarity list
                cur_sentiments.append(getTweetSentiment(tweet))
                # Append current tweets sentiments and polarity to the total for each product
                all_prod_tweets_sentiments.append(cur_sentiments)
        return {keys: values for keys, values in zip(product_names, all_prod_tweets_sentiments)}

# cleanTweet is used to clean a tweet from special characters and urls
# Accepts 1 argument
# @tweet: A raw tweet text
# ------------------------------------------------------------
def cleanTweet(tweet):
    '''
      Utility function to clean tweet text by removing links, special characters
      using simple regex statements.
    '''

    return ' '.join(re.sub("(@[A-Za-z0-9]+)|([^0-9A-Za-z ])  |(\w+:\/\/\S+)", " ", tweet).split())

# getPolarity returns the polarity of a tweet
# Accepts 1 argument
# @tweet: A raw tweet text
#------------------------------------------------------------
def getPolarity(tweet):
   analysis = TextBlob(cleanTweet(tweet))
   return analysis.sentiment.polarity


# getTweetSentiment returns the sentiment associated with a tweet
# Accepts 1 argument
# @tweet: A raw tweet text
# ------------------------------------------------------------
def getTweetSentiment(tweet):
    '''
    Utility function to classify sentiment of passed tweet
    using textblob's sentiment method
    '''

    # Return sentiment
    if getPolarity(tweet) > 0:
        return 'positive'
    elif getPolarity(tweet) == 0:
        return 'neutral'
    else:
        return 'negative'


# getTweets returns tweets that match a given query
# Accepts 3 arguments
# @query: Name of query to be executed
# @count: Number of tweets to be returned
# @api  : Twitter api config to use
# ------------------------------------------------------------
def getTweets(query, count, api):
    # empty list to store parsed tweets
    tweets = []
    # call twitter api to fetch tweets

    fetched_tweets = tweepy.Cursor(api.search,
                                   q=query,
                                   lang="en",
                                   since='2020-09-01').items(count)
    # parsing tweets one by one
    # print(fetched_tweets)

    for tweet in fetched_tweets:
        tweets.append(cleanTweet(tweet.text))
    return tweets

    # creating object of TwitterClient Class
    # calling function to get tweets

# get_consumer_key returns tweets consumer key
# ------------------------------------------------------------
def getConsumerKey():
    return "Hh2xwC7SWZaVOimVI3PKIvE4K"

# get_consumer_key returns tweets consumer key
# ------------------------------------------------------------
def getConsumerSecret():
    return "qPJqYQumNuu6tVgNwMniOJLXzk9AsLDes9EaclsivvkuCqVsZK"

# get_consumer_key returns tweets consumer key
# ------------------------------------------------------------
def getAccessToken():
    return "2418719884-vQ7Hz55Gl92ufgPqKjq5AH2vS7sCia18NYiDEhu"

# get_consumer_key returns tweets consumer key
# ------------------------------------------------------------
def getAccessTokenSecret():
    return "6r9OvYHd1TZVqZL7fIXRHo9EkmsQUQsjHkPmPEoNOM3bm"

# getTwitterAPI returns tweets a twitter API instance
# -----------------------------------------------------------
def getTwitterAPI():
    auth = tweepy.OAuthHandler(getConsumerKey(), getConsumerSecret())
    auth.set_access_token(getAccessToken(),getAccessTokenSecret())
    return tweepy.API(auth)

# getProdTweetsSentiments returns tweets a twitter API instance
# -----------------------------------------------------------
def getProdTweetsSentiments():
    all_prod_tweets_sentiments=[]
    for name in getProductsNames():
        cur_sentiments = []
        for tweet in getHealthProducts()[name]:
            # Add a sentiment to sentiments list
            cur_sentiments.append(getTweetSentiment(tweet))

        # Append current tweets sentiments and polarity to the total for each product
        all_prod_tweets_sentiments.append(cur_sentiments)
    return {keys: values for keys, values in zip(getProductsNames(), all_prod_tweets_sentiments)}

# getProdTweetsSentiments returns tweets a twitter API instance
# -----------------------------------------------------------
def getProdTweetsPolarity():
    all_prod_tweets_polarity=[]
    for name in getProductsNames():
        cur_tweets_polarity = []
        for tweet in getHealthProducts()[name]:
            # Add a polarity to tweets_polarity list
            cur_tweets_polarity.append(getPolarity(tweet))
        all_prod_tweets_polarity.append(cur_tweets_polarity)

    return {keys: values for keys, values in zip(getProductsNames(), all_prod_tweets_polarity)}
# getTwitterAPI returns tweets a twitter API instance
# -----------------------------------------------------------
def getHealthProducts():
    # Create a dictionary matching product names with their corresponding tweets
    return {keys: values for keys, values in zip(getProductsNames(), getProductsTweets())}

# getTwitterAPI returns tweets a twitter API instance
# -----------------------------------------------------------
def getProductsNames():
    return ["Omega-3", "Probiotic", "Cranberry Tea", "Luna bars", "Optimum Nutrition"]

# getTwitterAPI returns tweets a twitter API instance
# -----------------------------------------------------------
def getProductsTweets():
    product_tweets=[]
    for name in getQueryNames():
        product_tweets.append(getTweets(query=name + "-filter:retweets", count=5, api=getTwitterAPI()))
    return product_tweets

# getTwitterAPI returns tweets a twitter API instance
# -----------------------------------------------------------
def getQueryNames():
    return  ["brain Omega-3", "brain probiotic", "Cranberry Tea", "Luna bars", "Optimum Nutrition"]

# plotProdFeedback show feedback about the health product provided
# ----------------------------------------------------------------
def plotProdFeedback():
    # Set the size of the plot to 15 inches length and 8 width
    fig = plt.figure(figsize=(15, 8))
    for name in getProductsNames():
        plt.plot(getProdTweetsPolarity()[name], label='Tweets Polarity')
    plt.title('Feedback overtime')
    plt.xlabel('# of Tweets')
    plt.ylabel('Polarity from [-1,1]')
    plt.legend(getProductsNames())
    plt.show()
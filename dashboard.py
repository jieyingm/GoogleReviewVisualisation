import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from wordcloud import WordCloud
from collections import defaultdict, Counter
from textblob import TextBlob  # For sentiment intensity
from adjustText import adjust_text  # Helps avoid text overlap
import nltk
from nltk.corpus import stopwords
import dropbox
import datetime
import plotly.express as px

nltk.download("stopwords")
stop_words = set(stopwords.words("english"))

# Dropbox access token (Get from Dropbox App Console)
ACCESS_TOKEN = "sl.u.AFlZEbvHHvubZFH8aSIBmCP9mLCe0d3yEtc4rMhATfTdVK4lZrVjUoM23xUOg1o0hnsG_lne69Ji1OXlr87QvnUmuJ_SanoQbR2xzDc5d8nbvDaRbdchhDAVk6G7VrdZcBC3LCQ72mtaSfrKw2-T17I3ATveUkQ75P4wjyrqq9syKOSEvpK0xHkTUHCo4_kBu9nO65xCC7fSoksJbV_u3kEyKRg1yucfh95tcFIvw0IuElzItCjPRgt1-jKNzwX2hwODMEPUt1L89kiPs2Ft7Yw7IDwTE_EdP6apuuhhtD1K3I4_UW3CcNflnlUOVPhfCaflkf8Xqa4GXkjQodZfhbkPr7JAcGLGL1uwi7X9WXRcKa-LvMyRt4Wa7zEjEKH9MtP2raBooKlumSsJAdeWrZV031Qv1Ub8FFQjzHKqADxDfvVH88JOSLz27GcVMxT0cgFLqCFkhwfKm5kSG75wBqtEVQW4ecr0x0oipNTBHbhrAJLcuoACO4aLN5LIdemr6eM6rhaKWwOoVoz4gp1Zr8I3slrm4DpZtjC0nTcUU3GWLq7UNhmB3_e1VKVhoP719CMZSm99LGfOsBSmH9PJUeeFIAHwLd_Hew8W85FVDxaxec41QN9KNa3SfbBtjZmtOVKbwRQ3cE2Hr8cyU4GvgLISQDLpJd4xUxPAVruLAtgqem5DTkwu7SWr9pW0_2cw-uyP_n9ZpSongXfiA0GFouDkUKCSThDco3C1lvrBgdZaZR3PtAp0M78FUJ8uz4HXlse7zW66zbih_cWO-yR9UDLOcmZ7Wz2LwivIvBJQd_2YAx0w6_3LwjHfzEfMYHSF4lduYaSec3udGCnJjqBoxTKYD8zS1SL-jSFQct2TtaebgdIDQMqaV3qqn6KmUspzpWbpcnEiSiLmQKtfV1cnK7RpuJBiD59rVYxx2m00-TYiEy5Wzw5ckPMeS1HKOUTe14RmUzsX8bsqhnr-teOO8eD6s_JtggwUwUpqhxUVPQZA-2uPXhd7lfGnHd_bRff1Lm9odEwiT84Rp24JWtaTsFP0apjtNi6TFgddtnrdUx82uBhMoK_FAD7fjSTCvzibLhEjEAnImqOy0tIkcflVF4Z0l8oeHrixxlsBfH3a6m5O8aX3M0dl6MJuVu3CmZ4Fy0-9IUgUkWfzMfFnssyzZq5w3uuBKiIKHESDTVaZACPlDg1GluLorHBNXqeWFdfQpx4PDcnqg2u0bU7eHh35HkT-RLgK641j7a-V3G5qwMLjGYYYfqrg5Dp9254Qq87OHnBozz_tFtry1heOfhkfAC1ALyRg5IIkOkhydUhdq-oWtvXRSdeEw7jHodGuUFq-aMWlwVGO2lP_kwh_PllRW2P1s5Za0Xi2j3b2oIo1Pnjvuj17AfwARMuQqqeMPlbTDMc"
dbx = dropbox.Dropbox(ACCESS_TOKEN)

# Generate today's file name
today = datetime.datetime.today().strftime("%d-%m-%Y")
file_name = f"sentiment-test-{today}.xlsx"
dropbox_path = f"/UiPath/{file_name}"  # Update folder path

# Download file from Dropbox
local_file_path = file_name  # Save in the current working directory
try:
    dbx.files_download_to_file(local_file_path, dropbox_path)
    print(f"✅ File downloaded successfully: {local_file_path}")
except dropbox.exceptions.ApiError as e:
    print(f"❌ Error downloading file: {e}")

# Set file path variable for further use
file_path = local_file_path


# File path to sentiment analysis data
# file_path = "https://docs.google.com/spreadsheets/d/14YejDT2UB93Ah7Y0dUoZEXJlI62oDLuE/edit?usp=drive_link&ouid=109081502877770691586&rtpof=true&sd=true"

@st.cache_data
def get_shop_names():
    """Retrieve all sheet names from the Excel file (shop names)."""
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names  # Extract sheet names as shop names
    except FileNotFoundError:
        return []

@st.cache_data
def load_sentiment_data(sheet_name):
    """Load sentiment analysis data for the selected shop."""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Ensure 'Clean_Review' column is treated as strings and handle missing values
    df['Clean_Review'] = df['Clean_Review'].astype(str).replace('nan', '')
    
    return df

def categorize_complaints(negative_reviews):
    categories = {
        "Service": [
            "rude", "slow", "unfriendly", "ignored", "lambat", "lapar", "lewat", "unprofessional", "limited", "ignore",
            "sombong", "biadap", "teruk", "bodoh", "marah", "abaikan", "lazy", "careless", "attitude", "arrogant", "useless"
        ],
        "Food Quality": [
            "cold", "undercooked", "overcooked", "stale", "tasteless", "overate", "worse", "greasy", "terrible",
            "hampeh", "basi", "masin", "manis", "pahit", "hancur", "tawar", "lemau", "burnt", "hard", "raw", "soggy", "dry"
        ],
        "Pricing": [
            "expensive", "overprice", "costly", "bill", "mahal", "waste",
            "cekik", "melampau", "takberbaloi", "boros", "membazir", "scam", "ridiculous", "unfair", "inflated"
        ],
        "Cleanliness": [
            "dirty", "unclean", "hygiene", "smelly", "kotor", "nasty",
            "busuk", "berlendir", "berhabuk", "melekit", "berminyak", "bersepah", "lipas", "tikus",
            "filthy", "sticky", "messy", "dusty", "stinky", "gross"
        ],
        "Ambience": [
            "noisy", "loud", "dark", "bad", "bising", "tiny", "shout",
            "panas", "sempit", "sesak", "bau", "bingit", "terang",
            "cramped", "hot", "dim", "uncomfortable", "gloomy", "chaotic"
        ]
    }


    category_reviews = defaultdict(list)  # Dictionary to store reviews for each category

    for review in negative_reviews:
        review_lower = review.lower()
        for category, keywords in categories.items():
            if any(keyword in review_lower for keyword in keywords):
                category_reviews[category].append(review)  # Store the review

    return category_reviews

def get_sentiment_intensity(text):
    """Get sentiment intensity using TextBlob."""
    if isinstance(text, str) and text.strip():  # Check if text is a non-empty string
        blob = TextBlob(text)
        return blob.sentiment.polarity
    return 0  # Return neutral sentiment for empty or non-string values

def extract_frequent_words(reviews):
    """Extracts the most frequent words from reviews, filtering out stopwords."""
    words = " ".join(reviews).lower().split()
    words = [word for word in words if word.isalpha() and word not in stop_words]
    return Counter(words).most_common(10)  # Top 10 frequent words

# Streamlit UI
st.title(f"📊 Customer Sentiment Analysis Dashboard({today})")

shop_names = get_shop_names()

if shop_names:
    selected_shop = st.selectbox("🔍 Select a Shop:", shop_names)

    # Load sentiment data for the selected shop
    df = load_sentiment_data(selected_shop)

    st.subheader(f"📝 Customer Reviews for {selected_shop}")

    # Display full dataset
    # st.dataframe(df[["Name", "Date", "Review", "Sentiment"]])

    st.data_editor(
        df[["Name", "Date", "Review", "Sentiment"]],
        hide_index=True,
        column_config={
            "Name": st.column_config.TextColumn(width="small"),
            "Date": st.column_config.TextColumn(width="small"),
            "Review": st.column_config.TextColumn(width="medium"),  # Medium width for readability
            "Sentiment": st.column_config.TextColumn(width="small"),
        },
        height=500,  # Adjust height for better scrolling
        use_container_width=True,  # Ensures all columns fit on screen
    )

    # Sentiment summary
    st.subheader("📊 Sentiment Summary")

    # Count sentiment occurrences
    sentiment_counts = df["Sentiment"].value_counts().reset_index()
    sentiment_counts.columns = ["Sentiment", "Count"]

    # Ensure Sentiment column is string type
    sentiment_counts["Sentiment"] = sentiment_counts["Sentiment"].astype(str)

    # Define color mapping (green → red gradient)
    sentiment_colors = {
        "Very Positive": "#00FF00",  # Pastel Green
        "Positive": "#95FF66",       # Light Pastel Green
        "Neutral": "#FDFD96",        # Soft Peach
        "Negative": "#FFB347",       # Pastel Red-Orange
        "Very Negative": "#FF0000"   # Light Pink-Red
    }

    # Ensure only present sentiments are used in the color mapping
    filtered_colors = {k: v for k, v in sentiment_colors.items() if k in sentiment_counts["Sentiment"].values}

    # Create interactive Pie Chart with hover labels
    fig = px.pie(sentiment_counts, values="Count", names="Sentiment",
                color="Sentiment",  # Apply custom colors
                color_discrete_map=filtered_colors)  # Apply filtered colors

    # Update layout: Bigger fonts for legend & labels
    fig.update_layout(
        legend=dict(
            title="Sentiment", 
            x=1.05, y=1, 
            font=dict(size=20)  # Bigger font for legend
        ),
        margin=dict(l=40, r=160, t=40, b=40),  # Adjust margins for better spacing
        font=dict(size=14)  # Increase overall font size
    )

    # Update text inside the pie chart (percentages)
    fig.update_traces(
        textinfo='percent',  # Show both percentage and label
        textfont_size=16  # Bigger font for labels inside pie chart
    )

    # Display the chart in Streamlit
    st.plotly_chart(fig)



    # Word Cloud for Reviews
    st.subheader("☁️ Word Cloud of Reviews")
    all_reviews = " ".join(df["Clean_Review"].astype(str))  # Ensure all reviews are strings
    wordcloud = WordCloud(width=800, height=400, background_color="white").generate(all_reviews)
    
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.imshow(wordcloud, interpolation="bilinear")
    ax.axis("off")
    st.pyplot(fig)

    # 🔍 **New: Generate Word Cloud Insights**
    st.subheader("📢 Insights from Word Cloud")

    positive_reviews = df[df["Sentiment"].isin(["Positive", "Very Positive"])]["Clean_Review"]
    negative_reviews = df[df["Sentiment"] == "Negative"]["Clean_Review"]

    if not positive_reviews.empty:
        top_positive_words = extract_frequent_words(positive_reviews)
        st.success(f"✨ Customers **love**: {', '.join([word for word, _ in top_positive_words])}")

    if not negative_reviews.empty:
        top_negative_words = extract_frequent_words(negative_reviews)
        st.error(f"⚠️ Frequent **complaints** about: {', '.join([word for word, _ in top_negative_words])}")

    # 🚀 **Additional Insights**: Detect common topics
    common_topics = ["service", "food", "price", "cleanliness", "ambience"]
    detected_topics = [topic for topic in common_topics if topic in all_reviews]

    if detected_topics:
        st.info(f"💡 Key topics discussed: {', '.join(detected_topics)}")
    else:
        st.info("✅ No major topics detected.")

    # Complaint Cause Detection
    st.subheader("🚨 Complaint Cause Detection")
    negative_reviews = df[df["Sentiment"].isin(["Negative", "Neutral"])]["Clean_Review"]

    if not negative_reviews.empty:
        complaint_causes = categorize_complaints(negative_reviews)

        if complaint_causes:
            st.write("**Common Complaint Causes:**")
            category_counts = {category: len(reviews) for category, reviews in complaint_causes.items()}
            st.bar_chart(pd.Series(category_counts))

            # Recommended Improvements
            st.subheader("📢 Recommended Improvements")
            for category, reviews in complaint_causes.items():
                if category == "Service":
                    st.warning(f"💡 Improve customer service: {len(reviews)} complaints about service quality.")
                elif category == "Food Quality":
                    st.warning(f"💡 Improve food preparation: {len(reviews)} complaints about food quality.")
                elif category == "Pricing":
                    st.warning(f"💡 Consider promotions: {len(reviews)} complaints about pricing.")
                elif category == "Cleanliness":
                    st.warning(f"💡 Improve hygiene: {len(reviews)} complaints about cleanliness.")
                elif category == "Ambience":
                    st.warning(f"💡 Adjust atmosphere: {len(reviews)} complaints about ambience.")

                # Display related reviews
                with st.expander(f"📢 Read {len(reviews)} reviews about {category} issues"):
                    for review in reviews:
                        st.write(f"- {review}")

        else:
            st.success("✅ No major complaints detected! Keep up the good work.")

    else:
        st.success("✅ No major complaints detected! Keep up the good work.")

    # Separate Filter Reviews by Sentiment
    st.subheader("📌 Filter Reviews by Sentiment")

    # Let users choose a sentiment type
    selected_sentiment = st.radio("Select Sentiment Type:", df["Sentiment"].unique(), horizontal=True)

    # Filter reviews based on selection
    filtered_reviews = df[df["Sentiment"] == selected_sentiment]

    # Display the total count of selected reviews
    st.write(f"**Showing {len(filtered_reviews)} reviews for '{selected_sentiment}' sentiment:**")

    if not filtered_reviews.empty:
        # If more than 6 reviews, make it scrollable
        container = st.container()

        if len(filtered_reviews) > 6:
            with st.expander(f"🔍 View all {len(filtered_reviews)} reviews for '{selected_sentiment}'", expanded=True):
                container = st.container()

        # Display each review with ultra-compact spacing
        with container:
            for _, row in filtered_reviews.iterrows():
                st.markdown(f"**{row['Name']}** ({row['Date']})")  
                st.markdown(f"*{row['Review']}*", unsafe_allow_html=True)  # Italicized review for a sleek look
                st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)  # Ultra-thin divider
    else:
        st.info("No reviews found for this sentiment.")


else:
    st.warning("⚠️ No sentiment analysis data found. Waiting for UiPath to generate results.")

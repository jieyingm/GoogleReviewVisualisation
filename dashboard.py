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
ACCESS_TOKEN = "sl.u.AFnAEoUWfLnW9eZolfbCBReRo20_8favnoSHzWX2BYRGhrUVwzrpueggI9ZuayvleU1ArU9mqTUo62652x2y2c7MUxeW-Rnnai7G2leWZJvH0aH_H0ccRSZjnfrY_styQQg4Cf5xkY2WkHId45rfLQ5VaD9IGeq_f8g5SWamDrw5YzR2gPq8aDgj7SfLqzxTwqP9TIEv-ns3bXypOnssM-LeqOoENajrKCfK2SHUf0Bi7e2dhajGMOh-ac1WWydu6UYitGq1jtHDyra9LwmiIOA60NsMoHt3uBtmcZnIyfdG0U9xp4mRKEkb9HHyfPPlFYTDdocGCasa0R2OkbuEK-eKZ7W46JWWV8HL_U7mi28RlzCK87uBsGPB6qGmjNK9MCgx1zh7EjKsRH6PGmUV_BTtwUAx2Xx23FLeuWQ5YAGX1qyuHXZ5fNgbiq7qFR_nAd2CjsU0DCoSXDrckYmoGEFZvW2qHRe53EYnSCfFOqlS9syhqa6EdXbnldLhxSwU5mcTUsIN98UQzi9UzaD2-YO5e5DNHVD72d5fOvTQLwoYAmWRhWEW2f6gQ1BogAHssb_SQscQ9zStolwu6jfqgt5H18UzHQipQthWehHeF7jNeLCdXlLIlZguLpVFev9sP_WBurGpn_cSfz6jMgn9qSry2YEc_MX3CW5uI9uB9EWt_UBYu7ioYQVWXhsQC5HCkNaBBqoPwFhhLwDrMdME-6EnT50I2isJhwheSGC8cRO0b-yCmbLTjVDMniA5Gs9jVLG7_u_RcqyGBFFnwGiIvZMpSuzl5p9TaqESatcU_wWFGLGm3sm-Y3J10V53I1NYxVsyCetRaZAgF1z-Y9Z-Ll3Q-Bt-I0CSc75tVFbKQUJly27G2cd7TJ4W6S7OT4vLHoClhe5O2mUmfJs7LVHW0kMk1sftrJ2cyB6rxQZzOZ-f3UJ3AD8wSFDelzJ1F9vlAiJwXJTjAlL3aRB_ENwuMKg8OqoINAFMdCeB6n8NC_-j9_j56TZ_QllbJlbqOMMkW4fZ5i-tIN1w1SFZNVZfmgMV3MZF4rIutp_sqbJuycZ6f0ae5roNelduYsKJhQla8PRH0LfRt4WRFJvNOjBRVJT5DQlQYrev9y2jP8onlc7cPPSkuTLsUsVxKXZuQ4SXxqJRDeyQY0bphJaHj27jBmuR0peCQicdeuTYsSozPjYE2gTk-Ajhc8D3ymhJodGWld8rPDIHX6kHy0SXZo3Acc-04UWBpaXcBU-663FjJgkwuzTTyoZxCEc0LqSkon_BkapFbW4vngQTLc3YILEJJL6v_TSG-ajud3li7GOqSbckRwbGMxDqW26IUDW4sykypjOnsuEuFhbsoMWWSRYsbFiJLltNOxa144ROxZD107LjVTL4RUhIki6GeUkwgmwP69s"
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
st.title("📊 Customer Sentiment Analysis Dashboard")

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
        "Very Positive": "#77DD76",  # Pastel Green
        "Positive": "#BDE7BD",       # Light Pastel Green
        "Neutral": "#FDFD96",        # Soft Peach
        "Negative": "#FFB347",       # Pastel Red-Orange
        "Very Negative": "#FF6961"   # Light Pink-Red
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

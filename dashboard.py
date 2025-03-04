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

nltk.download("stopwords")
stop_words = set(stopwords.words("english"))

# File path to sentiment analysis data
file_path = "SentimentAnalysis_data.xlsx"

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
    """Categorizes negative reviews into key complaint causes and stores related reviews."""
    categories = {
        "Service": ["rude", "slow", "unfriendly", "ignored", "lambat", "lapar", "lewat"],
        "Food Quality": ["cold", "undercooked", "overcooked", "stale", "tasteless"],
        "Pricing": ["expensive", "overpriced", "costly", "bill", "mahal"],
        "Cleanliness": ["dirty", "unclean", "hygiene", "smelly", "kotor"],
        "Ambience": ["noisy", "loud", "dark", "bad atmosphere", "ambience", "environment", "bising"]
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
    sentiment_counts = df["Sentiment"].value_counts()

    # Improved Pie Chart
    fig, ax = plt.subplots(figsize=(6, 6))
    colors = sns.color_palette("pastel")  # Softer color palette

    # Create Pie Chart
    wedges, texts, autotexts = ax.pie(
        sentiment_counts,
        labels=None,  # Remove direct labels from slices
        autopct='%1.1f%%',
        startangle=90,
        colors=colors,
        wedgeprops={'edgecolor': 'black'},
        pctdistance=1.15,  # Move percentages outward
        labeldistance=1.3  # Ensure percentages do not overlap
    )

    # Adjust percentage label styles
    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_color("black")
        autotext.set_fontweight("bold")

    # Use adjustText to move labels slightly if they overlap
    texts = autotexts
    adjust_text(texts, expand=(1.2, 1.5), ax=ax)

    # Add a legend outside the chart
    ax.legend(wedges, sentiment_counts.index, title="Sentiment", loc="center left", bbox_to_anchor=(1, 0.5))

    ax.set_title("Sentiment Distribution", fontsize=14, fontweight='bold')

    st.pyplot(fig)


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
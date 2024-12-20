# TOPSIS-Based Text Conversational Analysis

This project applies the **TOPSIS (Technique for Order Preference by Similarity to Ideal Solution)** method to analyze conversational data based on various preprocessed parameters. The goal is to identify the best-performing models for different emotions (e.g., Happy, Sad, Anger, Thinking) across multiple conversation domains.
## Authors

- [@Guryansh](https://www.github.com/Guryansh)

## Features

- **TOPSIS Methodology**: Normalizes and weights multiple parameters to calculate relative closeness scores.
- **Emotion Analysis**: Separates data by emotion and identifies the best-performing model for each.
- **Excel Output**: Generates an Excel file with separate sheets for each emotion and a summary sheet of best models.
- **Custom Parameters**: Evaluates conversations based on the following parameters:
  - **Sentiment Score**: Range (-1 to 1), measures overall sentiment.
  - **Engagement**: Reply count, normalized.
  - **Response Time**: Time taken to respond, normalized.
  - **Clarity**: Flesch-Kincaid readability score.
  - **Relevance**: Cosine similarity score between user query and response.
  - **User Satisfaction Score**: Subjective feedback score.

## Workflow

The workflow of the analysis can be summarized as follows:

| **Step**               | **Process**                                                                                             |
|------------------------|---------------------------------------------------------------------------------------------------------|
| **Step 1: Data Input** | Load conversational data with emotions, domains, models, and evaluation parameters (e.g., P1 to P6).    |
| **Step 2: Preprocessing** | Normalize the parameters using Min-Max scaling for consistency.                                       |
| **Step 3: TOPSIS**     | Apply the TOPSIS method to calculate scores and ranks for each model in a given emotion-domain subset.   |
| **Step 4: Best Model** | Identify the best-performing model for each domain and overall for each emotion.                        |
| **Step 5: Export**     | Save results in an Excel file with separate sheets for each emotion and a summary sheet.                |

## Results Visualization

### Emotion-Specific Analysis

#### Happy  
![Happy Analysis](https://github.com/user-attachments/assets/f6a57e82-0cf1-40ca-a531-addf64bc056d)

#### Sad  
![Sad Analysis](https://github.com/user-attachments/assets/fdf070ad-d282-4e0d-8a9c-b0aa6adaf229)

#### Anger  
![Anger Analysis](https://github.com/user-attachments/assets/6d6205f5-94aa-43f5-aa0e-0d5226b5428f)

#### Thinking  
![Thinking Analysis](https://github.com/user-attachments/assets/83af3c76-9a52-4222-9a90-0faacf225e89)

### Overall Summary  
![Overall Results](https://github.com/user-attachments/assets/3b6e1412-4b07-429a-9f27-f0c62443eaec)

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)

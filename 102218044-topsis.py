import numpy as np
import pandas as pd
from sklearn.preprocessing import MinMaxScaler
from statistics import mode

data = {
    'Emotion': ['Happy'] * 16 + ['Sad'] * 16 + ['Anger'] * 16 + ['Think'] * 16,
    'Conversation Domain': ['C1', 'C1', 'C1', 'C1', 'C2', 'C2', 'C2', 'C2', 'C3', 'C3', 'C3', 'C3', 'C4', 'C4', 'C4',
                            'C4'] * 4,
    'Model': ['M1', 'M2', 'M3', 'M4'] * 16,
    'P1': np.random.uniform(-10, 10, 64),
    'P2': np.random.uniform(-10, 10, 64),
    'P3': np.random.uniform(-10, 10, 64),
    'P4': np.random.uniform(-10, 10, 64),
    'P5': np.random.uniform(-10, 10, 64),
    'P6': np.random.uniform(-10, 10, 64),
}

df = pd.DataFrame(data)


def apply_topsis(df_subset):
    criteria_cols = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6']

    scaler = MinMaxScaler()
    normalized_data = scaler.fit_transform(df_subset[criteria_cols])

    weights = np.array([1 / 6] * 6)

    ideal_solution = np.max(normalized_data, axis=0)
    negative_ideal_solution = np.min(normalized_data, axis=0)

    distance_to_ideal = np.sqrt(np.sum(((normalized_data - ideal_solution) ** 2) * weights, axis=1))
    distance_to_negative_ideal = np.sqrt(np.sum(((normalized_data - negative_ideal_solution) ** 2) * weights, axis=1))

    relative_closeness = distance_to_negative_ideal / (distance_to_ideal + distance_to_negative_ideal)

    return relative_closeness


best_models_per_emotion = {}

for emotion in df['Emotion'].unique():
    emotion_data = df[df['Emotion'] == emotion]

    topsis_scores = []
    ranks = []
    best_models = []

    for conversation_domain in emotion_data['Conversation Domain'].unique():
        subset = emotion_data[emotion_data['Conversation Domain'] == conversation_domain].copy()
        subset['TOPSIS Score'] = apply_topsis(subset)
        subset['Rank'] = subset['TOPSIS Score'].rank(ascending=False)

        best_model = subset.loc[subset['Rank'] == 1, 'Model'].values[0]
        best_models.append(best_model)

        topsis_scores.append(subset['TOPSIS Score'])
        ranks.append(subset['Rank'])

    overall_best_model = mode(best_models)
    best_models_per_emotion[emotion] = {'Best Model': overall_best_model}

    emotion_data['TOPSIS Score'] = np.concatenate(topsis_scores)
    emotion_data['Rank'] = np.concatenate(ranks)

    with pd.ExcelWriter("TOPSIS_Results_Emotions_Separate_Sheets.xlsx", mode='a', engine='openpyxl') as writer:
        emotion_data.to_excel(writer, sheet_name=emotion, index=False)

    print(f"{emotion}: Best Model = {overall_best_model}")

summary_df = pd.DataFrame.from_dict(best_models_per_emotion, orient='index').reset_index()
summary_df.rename(columns={'index': 'Emotion'}, inplace=True)

with pd.ExcelWriter("TOPSIS_Results_Emotions_Separate_Sheets.xlsx", mode='a', engine='openpyxl') as writer:
    summary_df.to_excel(writer, sheet_name="Best Models Summary", index=False)

print("Excel file with separate sheets for each emotion generated successfully!")

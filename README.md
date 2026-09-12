# Wasserstein Similarity Study

A Streamlit-based user study for investigating whether people can visually identify which of two displayed distributions is more similar to a reference distribution, using Wasserstein distance as the mathematical comparison criterion.

The study presents three 1D kernel-support distributions for each question:

- **Option A**
- **Reference distribution Q**
- **Option B**

Participants select the option that appears more similar to Q. After responding, the application reveals whether the selection corresponds to the option with the smaller Wasserstein distance and shows both Wasserstein-distance values.

## Live application

Open the deployed Streamlit application here:

> **[https://wasserstein-similarity-study.streamlit.app/](https://wasserstein-similarity-study.streamlit.app/)**

## Project structure

```text
wasserstein_similarity_study/
├── app.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
├── images/
│   ├── example/
│   │   └── wasserstein_example.png
│   ├── Q1/
│   │   ├── a.png
│   │   ├── b.png
│   │   └── t.png
│   ├── Q2/
│   │   ├── a.png
│   │   ├── b.png
│   │   └── t.png
│   └── ...
│       └── Q20/
│           ├── a.png
│           ├── b.png
│           └── t.png
├── create_wasserstein_example.py
├── study1_analysis.py
└── participant_results/
```

## Question-image convention

Each question uses three image files:

| File | Meaning |
|---|---|
| `images/Qi/a.png` | Option A distribution |
| `images/Qi/t.png` | Reference distribution Q |
| `images/Qi/b.png` | Option B distribution |

For example, Question 1 requires:

```text
images/Q1/a.png
images/Q1/t.png
images/Q1/b.png
```

Question 20 requires:

```text
images/Q20/a.png
images/Q20/t.png
images/Q20/b.png
```

The file paths are generated automatically by the application:

```python
QUESTIONS = [
    {
        "target_image": f"images/Q{i}/t.png",
        "option_a_image": f"images/Q{i}/a.png",
        "option_b_image": f"images/Q{i}/b.png",
        "correct_option": CORRECT_OPTIONS[i - 1],
        "distance_option_a": DISTANCES[i - 1],
        "distance_option_b": DISTANCES[i - 1],[3]
    }
    for i in range(1, 21)
]
```

## Local installation

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Start the application

```bash
streamlit run app.py
```

Streamlit will print a local address, usually:

```text
http://localhost:8501
```

Open that link in a web browser.

## Requirements

The application uses the following main Python packages:

```text
streamlit
pandas
openpyxl
matplotlib
numpy
```

A typical `requirements.txt` is:

```text
streamlit>=1.35.0
pandas>=2.0.0
openpyxl>=3.1.0
matplotlib>=3.7.0
numpy>=1.24.0
```

If you only run the Streamlit application and do not generate instructional plots or run analysis scripts, some packages may not be required. Keeping all listed packages makes the full repository reproducible.

## Configuring the study

The core study configuration is near the top of `app.py`.

### Correct options

Set the answer key in `CORRECT_OPTIONS`:

```python
CORRECT_OPTIONS = [
    "B",  # Q1
    "A",  # Q2
    "B",  # Q3
    # Continue until Q20
]
```

Each entry corresponds to one question. The first entry is used for Q1, the second for Q2, and so on.

### Wasserstein distances

Add the two distances for every question in `DISTANCES`:

```python
DISTANCES = [
    (8.000, 2.279),  # Q1: W(A, Q), W(B, Q)
    (0.111, 8.000),  # Q2: W(A, Q), W(B, Q)
    (6.000, 0.200),  # Q3: W(A, Q), W(B, Q)
]
```

Each tuple follows this exact order:

```python
(Wasserstein distance from Option A to Q,
 Wasserstein distance from Option B to Q)
```

For Question 1:

```python
(8.000, 2.279)
```

means:

\[
W(A, Q) = 8.000
\]

\[
W(B, Q) = 2.279
\]

Therefore, Option B is the correct answer for Q1 because it has the smaller distance.

### Number of questions

The current app uses 20 questions:

```python
for i in range(1, 21)
```

If you change the number of questions, update all of the following consistently:

- Number of image folders.
- Number of entries in `CORRECT_OPTIONS`.
- Number of entries in `DISTANCES`.
- The `range(1, 21)` expression.

For example, for 10 questions:

```python
for i in range(1, 11)
```

## Creating the instructional example plot

The repository includes a script that generates the example diagram used in the introductory slide:

```bash
python create_wasserstein_example.py
```

It creates:

```text
images/example/wasserstein_example.png
```

The figure shows:

- A compact Option A support distribution.
- A similar compact reference distribution Q.
- A wider and more distant Option B support distribution.
- Small vertical jitter for visual clarity.

The script also prints the two example Wasserstein distances. Copy those values into `app.py`:

```python
EXAMPLE_DISTANCE_A_TO_Q = 0.3000
EXAMPLE_DISTANCE_B_TO_Q = 2.9597
```

## Results

When a participant completes the study, the application provides:

- Final score.
- Accuracy percentage.
- Per-question answer summary.
- Correct option.
- Correct or incorrect result.
- \(W(A, Q)\).
- \(W(B, Q)\).
- A **Download Results** button that downloads the displayed result table as CSV.

The downloaded CSV is named:

```text
wasserstein_similarity_results.csv
```

It contains columns similar to:

```csv
Question,Your Answer,Correct Answer,Result,"W(A, Q)","W(B, Q)"
Q1,Option B,Option B,Correct,8.0,2.279
Q2,Option A,Option A,Correct,0.111,8.0
```

## Analysis of downloaded results

Use `study1_analysis.py` to analyze downloaded participant CSV files.

Expected analysis directory structure:

```text
Result_analysis/
├── study1_analysis.py
├── answer_key.csv
└── participant_results/
    ├── participant_01.csv
    ├── participant_02.csv
    ├── participant_03.csv
    └── participant_04.csv
```

The answer key must include:

```csv
Question,Correct Answer
Q1,Option B
Q2,Option A
Q3,Option B
```

The script supports any number of participant files and generates:

```text
study_analysis/
├── all_responses_merged.csv
├── question_summary.csv
├── response_matrix.csv
├── response_matrix.xlsx
├── pairwise_agreement.csv
├── participant_accuracy.png
├── question_accuracy.png
├── response_heatmap.png
└── pairwise_agreement_heatmap.png
```

Run the analysis with:

```bash
python study1_analysis.py
```


### Analysis outputs


| Output | Description |
|---|---|
| `all_responses_merged.csv` | All participant responses merged with the answer key and Wasserstein-distance information for analysis |
| `question_summary.csv` | Per-question counts for selected A, selected B, correct responses, majority choice, both Wasserstein distances, and the Wasserstein-distance gap |
| `response_matrix.csv` | Matrix of participant A/B selections for every question, including the correct option |
| `response_matrix.xlsx` | Formatted response matrix with correct selections highlighted in light green |
| `pairwise_agreement.csv` | Percentage agreement between every pair of participants |
| `participant_accuracy.png` | Percentage of responses matching the lower-Wasserstein option for each participant |
| `question_accuracy.png` | Percentage of participants selecting the lower-Wasserstein option for each question |
| `response_heatmap.png` | Participant response matrix with correct selections highlighted in light green |
| `pairwise_agreement_heatmap.png` | Heatmap showing pairwise agreement between participants |


## Deployment with Streamlit Community Cloud

To create a shareable application link:

1. Push this repository to GitHub.
2. Go to [Streamlit Community Cloud](https://share.streamlit.io/).
3. Sign in with GitHub.
4. Select **Create app**.
5. Choose this repository and the desired branch, usually `main`.
6. Set the main file path to:

   ```text
   app.py
   ```

7. Click **Deploy**.
8. Copy and share the generated `streamlit.app` link.

For a public repository, the standard GitHub OAuth permissions are generally sufficient to deploy. A `requirements.txt` should be placed in the repository root or alongside the Streamlit entry-point file so Community Cloud can install dependencies.

## Theme configuration

The application is configured for light mode. Create:

```text
.streamlit/config.toml
```

with:

```toml
[theme]
base = "light"
```

## Reproducibility notes

- Keep all question images under version control.
- Use stable question IDs, such as Q1 to Q20.
- Keep the order of `CORRECT_OPTIONS` and `DISTANCES` aligned with question numbering.
- Record the code version or Git commit hash used for a study run.
- Store downloaded participant CSV files using unique participant filenames.

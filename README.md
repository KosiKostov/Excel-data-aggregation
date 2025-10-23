# Next Word Prediction using LSTM

A deep learning-based text generation system trained on Sherlock Holmes stories to predict the next words in a sequence.

## Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Model Architecture](#model-architecture)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Model Files](#model-files)
- [Limitations](#limitations)

## Overview

This project uses a Long Short-Term Memory (LSTM) neural network to learn patterns from text and predict subsequent words. The model is trained on Sherlock Holmes stories and provides a simple GUI for interactive text prediction.

## Project Structure

```
.
├── train_v3.py                      # Training script
├── predict_v3.py                    # Prediction GUI application
├── sherlock-holm.es_stories_plain-text_advs.txt  # Training data
├── next_word_model_v3_simple.h5     # Trained Keras model
├── tokenizer_v3_simple.pkl          # Tokenizer for text preprocessing
├── model_metadata_v3_simple.pkl     # Model configuration metadata
└── README.md                        # This file
```

## Requirements


```bash
pip install numpy tensorflow tkinter pickle-mixin
```

- **Python**: 3.7+
- **TensorFlow**: 2.x
- **NumPy**: For numerical operations
- **Tkinter**: For GUI 
- **Pickle**: For serialization 

## Model Architecture

### Network Structure

- **Embedding Layer**: 100-dimensional word embeddings
- **LSTM Layer**: 150 units for sequence processing
- **Dense Output Layer**: Softmax activation for word probability distribution

### Training Configuration

- **Loss Function**: Categorical Cross-Entropy
- **Optimizer**: Adam
- **Epochs**: 100
- **Input**: N-gram sequences from text
- **Output**: Next word prediction (one-hot encoded)

## Usage

### 1. Training the Model

Run the training script to create a new model:

```bash
python train_v3.py
```

**What it does:**
- Loads the Sherlock Holmes text file
- Creates a tokenizer and vocabulary
- Generates n-gram sequences for training
- Trains an LSTM model for 100 epochs
- Saves the model, tokenizer, and metadata

**Output files:**
- `next_word_model_v3_simple.h5` - Trained model
- `tokenizer_v3_simple.pkl` - Fitted tokenizer
- `model_metadata_v3_simple.pkl` - Model configuration


### 2. Running the Prediction GUI

Launch the graphical interface:

```bash
python predict_v3.py
```

**GUI Features:**
- **Text Input Area**: Enter your seed text
- **Word Count Selector**: Select number of words to predict (1-10)
- **Predict Button**: Generate predictions
- **Clear Button**: Reset all fields
- **Result Area**: Displays the complete text with predictions


## How It Works

### Training Process

1. **Data Loading**: Reads the Sherlock Holmes text file
2. **Tokenization**: Converts words to numerical indices
3. **Sequence Generation**: Creates n-gram sequences from each line
   - For the sentence "the cat sat", it creates:
     - [the, cat]
     - [the, cat, sat]
4. **Padding**: Sequences are padded to uniform length
5. **Data Preparation**: 
   - X: All words except the last
   - y: The last word (target)
6. **Training**: Model learns to predict the next word given the sequence



## Model Files

### next_word_model_v3_simple.h5
The trained Keras model containing:
- Network architecture
- Learned weights and biases
- Compilation configuration

### tokenizer_v3_simple.pkl
Serialized tokenizer containing:
- Word-to-index mappings
- Index-to-word mappings
- Vocabulary size

### model_metadata_v3_simple.pkl
Configuration metadata:
- `max_sequence_len`: Maximum input sequence length
- `total_words`: Vocabulary size

## Limitations

- Predictions are based solely on the training corpus (Sherlock Holmes stories)
- Model generates one word at a time greedily (no beam search)
- No temperature control for prediction diversity
- Limited context window based on max sequence length
- May generate repetitive or nonsensical sequences

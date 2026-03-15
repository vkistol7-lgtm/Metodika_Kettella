
import streamlit as st
import pandas as pd
import io
import docx

def parse_docx(uploaded_file):
    """Парсит docx файл и возвращает список словарей с вопросами и вариантами."""
    doc = docx.Document(uploaded_file)
    questions = []
    current_text = []
    options = {}
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        # Определение вариантов ответа (учитывается кириллица и латиница)
        text_lower = text.lower()
        if text_lower.startswith(('a)', 'а)')):
            options['a'] = text
        elif text_lower.startswith(('b)', 'в)')):
            options['b'] = text
        elif text_lower.startswith(('c)', 'с)')):
            options['c'] = text
        else:
            if options:
                # Сохранение предыдущего вопроса при обнаружении нового
                questions.append({
                    'text': '\n'.join(current_text),
                    'options': options
                })
                current_text = [text]
                options = {}
            else:
                current_text.append(text)
                
    # Добавление последнего вопроса
    if current_text and options:
        questions.append({
            'text': '\n'.join(current_text),
            'options': options
        })
    return questions

def generate_excel(answers, total_questions):
    """Формирует DataFrame и конвертирует в байтовый поток Excel."""
    # Создание пустой таблицы с нужными столбцами
    df = pd.DataFrame(index=range(1, total_questions + 1), columns=['a', 'b', 'c'])
    df.index.name = '№'
    
    # Заполнение таблицы ответами
    for q_idx, ans_letter in answers.items():
        q_num = q_idx + 1
        if ans_letter in df.columns:
            df.at[q_num, ans_letter] = ans_letter

    # Запись в буфер
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Бланк ответов')
    return output.getvalue()

def main():
    st.set_page_config(page_title="Тестирование", layout="centered")

    # Инициализация переменных состояния
    if 'questions' not in st.session_state:
        st.session_state.questions = []
    if 'current_q' not in st.session_state:
        st.session_state.current_q = 0
    if 'answers' not in st.session_state:
        st.session_state.answers = {}
    if 'test_finished' not in st.session_state:
        st.session_state.test_finished = False

    st.title("Система тестирования")

    # Блок загрузки файла с вопросами
    if not st.session_state.questions:
        st.info("Загрузите файл с вопросами в формате .docx для начала работы.")
        uploaded_file = st.file_uploader("Файл теста", type="docx")
        if uploaded_file is not None:
            try:
                st.session_state.questions = parse_docx(uploaded_file)
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка при обработке файла: {e}")
        return

    questions = st.session_state.questions
    total_q = len(questions)

    # Экран завершения теста
    if st.session_state.test_finished:
        st.success("Тестирование завершено. Скачайте бланк ответов.")
        
        excel_data = generate_excel(st.session_state.answers, total_q)
        st.download_button(
            label="Скачать бланк ответов (Excel)",
            data=excel_data,
            file_name="Бланк_ответов.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if st.button("Начать заново"):
            st.session_state.clear()
            st.rerun()
        return

    # Экран текущего вопроса
    current_idx = st.session_state.current_q
    q_data = questions[current_idx]

    st.progress((current_idx) / total_q)
    st.write(f"**Вопрос {current_idx + 1} из {total_q}**")
    st.write(q_data['text'])

    # Определение текущего сохраненного ответа
    current_answer = st.session_state.answers.get(current_idx, None)
    option_keys = list(q_data['options'].keys())
    
    # Индекс для radio button
    index = option_keys.index(current_answer) if current_answer in option_keys else 0

    selected_option = st.radio(
        "Выберите вариант:",
        option_keys,
        index=index,
        format_func=lambda x: q_data['options'][x]
    )

    # Кнопки навигации
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if current_idx > 0:
            if st.button("Назад"):
                st.session_state.answers[current_idx] = selected_option
                st.session_state.current_q -= 1
                st.rerun()

    with col3:
        if current_idx < total_q - 1:
            if st.button("Далее"):
                st.session_state.answers[current_idx] = selected_option
                st.session_state.current_q += 1
                st.rerun()
        else:
            if st.button("Завершить"):
                st.session_state.answers[current_idx] = selected_option
                st.session_state.test_finished = True
                st.rerun()

if __name__ == "__main__":
    main()
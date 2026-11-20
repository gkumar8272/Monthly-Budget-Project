import React from 'react';

const ExpenseTable = ({ expenses, deleteExpens }) => {

    return (
        <section className="expense-list-section">
            <div className='section-heading'>
                <div><p className='eyebrow'>ACTIVITY</p><h2>Recent transactions</h2></div>
                <span className='transaction-count'>{expenses.length} {expenses.length === 1 ? 'entry' : 'entries'}</span>
            </div>
            <div className="expense-list">
                {expenses.length === 0 ? <p className='empty-state'>Your transactions will appear here.</p> : expenses.map((expense, index) => (
                    <div key={index} className="expense-item">
                        <span className={`transaction-marker ${expense.amount > 0 ? 'income-marker' : ''}`}></span>
                        <div className="expense-description"><strong>{expense.text}</strong><span>{expense.amount > 0 ? 'Income' : 'Expense'}</span></div>
                        <div className={`expense-amount ${expense.amount > 0 ? 'income-amount' : ''}`}>{expense.amount > 0 ? '+' : '-'}₹{Math.abs(expense.amount).toLocaleString('en-IN')}</div>
                        <button className="delete-button" aria-label={`Delete ${expense.text}`} onClick={() => deleteExpens(expense._id)}>×</button>
                    </div>
                ))}
            </div>
        </section>
    );
};

export default ExpenseTable;

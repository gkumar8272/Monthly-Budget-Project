import React, { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom';
import { APIUrl, handleError, handleSuccess } from '../utils';
import { ToastContainer } from 'react-toastify';
import ExpenseTable from './ExpenseTable';
import ExpenseDetails from './ExpenseDetails';
import ExpenseForm from './ExpenseForm';

function Home() {
    const [loggedInUser, setLoggedInUser] = useState('');
    const [expenses, setExpenses] = useState([]);
    const [incomeAmt, setIncomeAmt] = useState(0);
    const [expenseAmt, setExpenseAmt] = useState(0);

    const navigate = useNavigate();

    useEffect(() => {
        setLoggedInUser(localStorage.getItem('loggedInUser'))
    }, [])

    const handleLogout = (e) => {
        localStorage.removeItem('token');
        localStorage.removeItem('loggedInUser');
        handleSuccess('User Loggedout');
        setTimeout(() => {
            navigate('/login');
        }, 1000)
    }
    useEffect(() => {
        const amounts = expenses.map(item => item.amount);
        const income = amounts.filter(item => item > 0)
            .reduce((acc, item) => (acc += item), 0);
        const exp = amounts.filter(item => item < 0)
            .reduce((acc, item) => (acc += item), 0) * -1;
        setIncomeAmt(income);
        setExpenseAmt(exp);
    }, [expenses])

    const deleteExpens = async (id) => {
        try {
            const url = `${APIUrl}/expenses/${id}`;
            const headers = {
                headers: {
                    'Authorization': localStorage.getItem('token')
                },
                method: "DELETE"
            }
            const response = await fetch(url, headers);
            if (response.status === 403) {
                localStorage.removeItem('token');
                navigate('/login');
                return
            }
            const result = await response.json();
            handleSuccess(result?.message)
            console.log('--result', result.data);
            setExpenses(result.data);
        } catch (err) {
            handleError(err);
        }
    }

    const addTransaction = async (data) => {
        try {
            const url = `${APIUrl}/expenses`;
            const headers = {
                headers: {
                    'Authorization': localStorage.getItem('token'),
                    'Content-Type': 'application/json'
                },
                method: "POST",
                body: JSON.stringify(data)
            }
            const response = await fetch(url, headers);
            if (response.status === 403) {
                localStorage.removeItem('token');
                navigate('/login');
                return
            }
            const result = await response.json();
            handleSuccess(result?.message)
            console.log('--result', result.data);
            setExpenses(result.data);
        } catch (err) {
            handleError(err);
        }
    }

    useEffect(() => {
        const fetchExpenses = async () => {
            try {
                const url = `${APIUrl}/expenses`;
                const headers = {
                    headers: {
                        'Authorization': localStorage.getItem('token')
                    }
                }
                const response = await fetch(url, headers);
                if (response.status === 403) {
                    localStorage.removeItem('token');
                    navigate('/login');
                    return
                }
                const result = await response.json();
                console.log('--result', result.data);
                setExpenses(result.data);
            } catch (err) {
                handleError(err);
            }
        }

        fetchExpenses()
    }, [navigate])

    return (
        <main className='dashboard-shell'>
            <div className='user-section'>
                <div>
                    <p className='eyebrow'>PERSONAL FINANCE</p>
                    <h1>Good to see you, {loggedInUser}</h1>
                    <p className='page-subtitle'>Keep your everyday spending clear and intentional.</p>
                </div>
                <button className='logout-button' onClick={handleLogout}>Log out</button>
            </div>
            <section className='dashboard-grid'>
                <div className='dashboard-main'>
                    <ExpenseDetails incomeAmt={incomeAmt} expenseAmt={expenseAmt} />
                    <ExpenseTable expenses={expenses} deleteExpens={deleteExpens} />
                </div>
                <aside className='dashboard-aside'>
                    <ExpenseForm addTransaction={addTransaction} />
                    <div className='tip-card'>
                        <span className='tip-mark'>i</span>
                        <div>
                            <h3>Small habit, big clarity</h3>
                            <p>Record purchases as they happen to keep your balance honest.</p>
                        </div>
                    </div>
                </aside>
            </section>
            <ToastContainer />
        </main>
    )
}

export default Home
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp_DialogueAddSignatories.ViewModel
{
    /// <summary>
    /// Класс ViewModelBase - предоставляет возможность всем ViewModel оповещать view об изменениях
    /// которые в них происходят 
    /// IDisposable - интерфейс восвобождения ресурсов
    /// </summary>
    public abstract class ViewModelBase : INotifyPropertyChanged, IDisposable
    {
        protected ViewModelBase()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Этот метод будет вызывать наши события
        /// </summary>
        /// <param name="propertyName"></param>
        public virtual void OnPropertyChanged(string propertyName)
        {
            //это практически копия 19 строки, 
            //это нужно когда реализуем этот метод в многопоточных приложениях
            //потому что в данном случае для каждого потока будет создаваться своя
            //локальная переменная handler и он будет взаимодействовать с данными которые находятся
            //в этой локальнйо переменной 
            //в противном же случаи потоки бы пользовались одним событием и если 1 поток его изменил
            //то другой поток может получить некоректные данные
            PropertyChangedEventHandler handler = this.PropertyChanged;

            if (handler != null)
            {
                handler.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public void Dispose()
        {
            this.OnDispose();
        }

        protected virtual void OnDispose()
        {
        }
    }
}



using System;
using System.Text;

namespace GeoGroup.ExcelExtension.External
{
    public static class WritingSumConverter
    {
        public static StringBuilder ToWriting(decimal value, Currency currency, StringBuilder result)
        {
            decimal integerPart = Math.Floor(value);
            uint fractionalPart = (uint)((value - integerPart) * 100);
            Number.ToWriting(integerPart, currency.ОсновнаяЕдиница, result);
            return AddCoins(fractionalPart, currency, result);
        }

        public static StringBuilder ToWriting(double value, Currency currency, StringBuilder result)
        {
            double integerPart = Math.Floor(value);
            uint fractionalPart = (uint)(value * 100) - (uint)(integerPart * 100);
            Number.Пропись(integerPart, currency.ОсновнаяЕдиница, result);
            return AddCoins(fractionalPart, currency, result);
        }

        private static StringBuilder AddCoins(uint value, Currency валюта, StringBuilder result)
        {
            result.Append(' ');
            result.Append(value.ToString("00"));
            result.Append(' ');
            result.Append(Number.Согласовать(валюта.ДробнаяЕдиница, value));
            return result;
        }

        public static string ToWriting(decimal n, Currency валюта)
        {
            return Number.ApplyCaps(ToWriting(n, валюта, new StringBuilder()), Заглавные.Нет);
        }

        public static string ToWriting(decimal n, Currency валюта, Заглавные заглавные)
        {
            return Number.ApplyCaps(ToWriting(n, валюта, new StringBuilder()), заглавные);
        }

        public static string ToWriting(double n, Currency валюта)
        {
            return Number.ApplyCaps(ToWriting(n, валюта, new StringBuilder()), Заглавные.Нет);
        }

        public static string ToWriting(double n, Currency валюта, Заглавные заглавные)
        {
            return Number.ApplyCaps(ToWriting(n, валюта, new StringBuilder()), заглавные);
        }
    }

    public static class Number
    {
        public static StringBuilder ToWriting(decimal число, IЕдиницаИзмерения еи, StringBuilder result)
        {
            string error = ПроверитьЧисло(число);
            if (error != null) throw new ArgumentException(error, nameof(число));

            if (число <= uint.MaxValue)
            {
                Пропись((uint)число, еи, result);
            }

            else if (число <= ulong.MaxValue)
            {
                Пропись((ulong)число, еи, result);
            }
            else
            {
                SpaceAppender mySb = new SpaceAppender(result);
                decimal div1000 = Math.Floor(число / 1000);
                ПрописьСтаршихКлассов(div1000, 0, mySb);
                ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
            }
            return result;
        }

        public static StringBuilder Пропись(double число, IЕдиницаИзмерения еи, StringBuilder result)
        {
            string error = ПроверитьЧисло(число);
            if (error != null) throw new ArgumentException(error, nameof(число));
            if (число <= uint.MaxValue)
            {
                Пропись((uint)число, еи, result);
            }

            else if (число <= ulong.MaxValue)
            {
                Пропись((ulong)число, еи, result);
            }
            else
            {
                SpaceAppender mySb = new SpaceAppender(result);
                double div1000 = Math.Floor(число / 1000);
                ПрописьСтаршихКлассов(div1000, 0, mySb);
                ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
            }
            return result;
        }

        public static StringBuilder Пропись(ulong число, IЕдиницаИзмерения еи, StringBuilder result)
        {
            if (число <= uint.MaxValue)
            {
                Пропись((uint)число, еи, result);
            }
            else
            {
                SpaceAppender mySb = new SpaceAppender(result);
                ulong div1000 = число / 1000;
                ПрописьСтаршихКлассов(div1000, 0, mySb);
                ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
            }
            return result;
        }

        public static StringBuilder Пропись(uint число, IЕдиницаИзмерения еи, StringBuilder result)
        {
            SpaceAppender mySb = new SpaceAppender(result);
            if (число == 0)
            {
                mySb.Append("ноль");
                mySb.Append(еи.РодМножественное);
            }
            else
            {
                uint div1000 = число / 1000;
                ПрописьСтаршихКлассов(div1000, 0, mySb);
                ПрописьКласса(число - div1000 * 1000, еи, mySb);
            }
            return result;
        }

        private static void ПрописьСтаршихКлассов(decimal число, int номерКласса, SpaceAppender sb)
        {
            if (число == 0) return; // конец рекурсии
            decimal div1000 = Math.Floor(число / 1000);
            ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);
            uint числоДо999 = (uint)(число - div1000 * 1000);
            if (числоДо999 == 0) return;
            ПрописьКласса(числоДо999, Классы[номерКласса], sb);
        }


        private static void ПрописьСтаршихКлассов(double число, int номерКласса, SpaceAppender sb)
        {
            if (число == 0) return;
            double div1000 = Math.Floor(число / 1000);
            ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);
            uint числоДо999 = (uint)(число - div1000 * 1000);
            if (числоДо999 == 0) return;
            ПрописьКласса(числоДо999, Классы[номерКласса], sb);
        }

        private static void ПрописьСтаршихКлассов(ulong число, int номерКласса, SpaceAppender sb)
        {
            if (число == 0) return; // конец рекурсии
            ulong div1000 = число / 1000;
            ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);
            uint числоДо999 = (uint)(число - div1000 * 1000);
            if (числоДо999 == 0) return;
            ПрописьКласса(числоДо999, Классы[номерКласса], sb);
        }

        private static void ПрописьСтаршихКлассов(uint число, int номерКласса, SpaceAppender sb)
        {
            if (число == 0) return;
            uint div1000 = число / 1000;
            ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);
            uint числоДо999 = число - div1000 * 1000;
            if (числоДо999 == 0) return;
            ПрописьКласса(числоДо999, Классы[номерКласса], sb);
        }

        private static void ПрописьКласса(uint числоДо999, IЕдиницаИзмерения класс, SpaceAppender sb)
        {
            uint числоЕдиниц = числоДо999 % 10;
            uint числоДесятков = (числоДо999 / 10) % 10;
            uint числоСотен = (числоДо999 / 100) % 10;

            sb.Append(Сотни[числоСотен]);
            if ((числоДо999 % 100) != 0)
            {
                Десятки[числоДесятков].Пропись(sb, числоЕдиниц, класс.РодЧисло);
            }
            sb.Append(Согласовать(класс, числоДо999));
        }

        public static string ПроверитьЧисло(decimal число)
        {
            if (число < 0)
                return "Число должно быть больше или равно нулю.";
            if (число != decimal.Floor(число))
                return "Число не должно содержать дробной части.";
            return null;
        }

        public static string ПроверитьЧисло(double число)
        {
            if (число < 0)
                return "Число должно быть больше или равно нулю.";
            if (число != Math.Floor(число))
                return "Число не должно содержать дробной части.";
            if (число > MaxDouble)
            {
                return "Число должно быть не больше " + MaxDouble + ".";
            }
            return null;
        }

        public static string Согласовать(IЕдиницаИзмерения единицаИзмерения, uint число)
        {
            uint числоЕдиниц = число % 10;
            uint числоДесятков = (число / 10) % 10;
            if (числоДесятков == 1) return единицаИзмерения.РодМножественное;
            switch (числоЕдиниц)
            {
                case 1: return единицаИзмерения.ИменЕдинственное;
                case 2: case 3: case 4: return единицаИзмерения.РодЕдинственное;
                default: return единицаИзмерения.РодМножественное;
            }
        }

        private static string ПрописьЦифры(uint цифра, РодЧисло род)
        {
            return Цифры[цифра].Пропись(род);
        }

        abstract class Цифра
        {
            public abstract string Пропись(РодЧисло род);
        }

        private class ЦифраИзменяющаясяПоРодам : Цифра, IИзменяетсяПоРодам
        {
            public ЦифраИзменяющаясяПоРодам(
                string мужской,
                string женский,
                string средний,
                string множественное)
            {
                Мужской = мужской;
                Женский = женский;
                Средний = средний;
                Множественное = множественное;
            }

            public ЦифраИзменяющаясяПоРодам(
                string единственное,
                string множественное)
                : this(единственное, единственное, единственное, множественное)
            {
            }

            public string Мужской { get; }
            public string Женский { get; }
            public string Средний { get; }
            public string Множественное { get; }

            public override string Пропись(РодЧисло род)
            {
                return род.ПолучитьФорму(this);
            }
        }

        private class ЦифраНеИзменяющаясяПоРодам : Цифра
        {
            public ЦифраНеИзменяющаясяПоРодам(string пропись)
            {
                _пропись = пропись;
            }

            private readonly string _пропись;
            public override string Пропись(РодЧисло род)
            {
                return _пропись;
            }
        }

        private static readonly Цифра[] Цифры = {
            null,
            new ЦифраИзменяющаясяПоРодам ("один", "одна", "одно", "одни"),
            new ЦифраИзменяющаясяПоРодам ("два", "две", "два", "двое"),
            new ЦифраИзменяющаясяПоРодам ("три", "трое"),
            new ЦифраИзменяющаясяПоРодам ("четыре", "четверо"),
            new ЦифраНеИзменяющаясяПоРодам ("пять"),
            new ЦифраНеИзменяющаясяПоРодам ("шесть"),
            new ЦифраНеИзменяющаясяПоРодам ("семь"),
            new ЦифраНеИзменяющаясяПоРодам ("восемь"),
            new ЦифраНеИзменяющаясяПоРодам ("девять"),
        };

        private static readonly Десяток[] Десятки = {
            new ПервыйДесяток (),
            new ВторойДесяток (),
            new ОбычныйДесяток ("двадцать"),
            new ОбычныйДесяток ("тридцать"),
            new ОбычныйДесяток ("сорок"),
            new ОбычныйДесяток ("пятьдесят"),
            new ОбычныйДесяток ("шестьдесят"),
            new ОбычныйДесяток ("семьдесят"),
            new ОбычныйДесяток ("восемьдесят"),
            new ОбычныйДесяток ("девяносто")
        };

        abstract class Десяток
        {
            public abstract void Пропись(SpaceAppender sb, uint числоЕдиниц, РодЧисло род);
        }

        private class ПервыйДесяток : Десяток
        {
            public override void Пропись(SpaceAppender sb, uint числоЕдиниц, РодЧисло род)
            {
                sb.Append(ПрописьЦифры(числоЕдиниц, род));
            }
        }

        private class ВторойДесяток : Десяток
        {
            private static readonly string[] WritingFromTenToTwenty = {
                "десять",
                "одиннадцать",
                "двенадцать",
                "тринадцать",
                "четырнадцать",
                "пятнадцать",
                "шестнадцать",
                "семнадцать",
                "восемнадцать",
                "девятнадцать"
            };

            public override void Пропись(SpaceAppender sb, uint числоЕдиниц, РодЧисло род)
            {
                sb.Append(WritingFromTenToTwenty[числоЕдиниц]);
            }
        }

        private class ОбычныйДесяток : Десяток
        {
            public ОбычныйДесяток(string названиеДесятка)
            {
                _названиеДесятка = названиеДесятка;
            }

            private readonly string _названиеДесятка;

            public override void Пропись(SpaceAppender sb, uint числоЕдиниц, РодЧисло род)
            {
                sb.Append(_названиеДесятка);
                if (числоЕдиниц == 0)
                {
                }
                else
                {
                    sb.Append(ПрописьЦифры(числоЕдиниц, род));
                }
            }
        }

        private static readonly string[] Сотни = {
            null,
            "сто",
            "двести",
            "триста",
            "четыреста",
            "пятьсот",
            "шестьсот",
            "семьсот",
            "восемьсот",
            "девятьсот"
        };

        private class КлассТысяч : IЕдиницаИзмерения
        {
            public string ИменЕдинственное => "тысяча";
            public string РодЕдинственное => "тысячи";
            public string РодМножественное => "тысяч";
            public РодЧисло РодЧисло => РодЧисло.Женский;
        }

        private class Класс : IЕдиницаИзмерения
        {
            public Класс(string начальнаяФорма)

            {
                ИменЕдинственное = начальнаяФорма;
            }

            public string ИменЕдинственное { get; }
            public string РодЕдинственное => ИменЕдинственное + "а";
            public string РодМножественное => ИменЕдинственное + "ов";
            public РодЧисло РодЧисло => РодЧисло.Мужской;
        }

        private static readonly IЕдиницаИзмерения[] Классы = {
            new КлассТысяч (),
            new Класс ("миллион"),
            new Класс ("миллиард"),
            new Класс ("триллион"),
            new Класс ("квадриллион"),
            new Класс ("квинтиллион"),
            new Класс ("секстиллион"),
            new Класс ("септиллион")
        };

        public static double MaxDouble
        {
            get
            {
                if (_maxDouble == 0)
                {
                    _maxDouble = CalcMaxDouble();
                }
                return _maxDouble;
            }
        }

        private static double _maxDouble;

        private static double CalcMaxDouble()
        {
            double max = Math.Pow(1000, Классы.Length + 1);
            double d = 1;
            while (max - d == max)
            {
                d *= 2;
            }
            return max - d;
        }

        private class SpaceAppender
        {
            public SpaceAppender(StringBuilder stringBuilder)
            {
                _stringBuilder = stringBuilder;
            }

            private readonly StringBuilder _stringBuilder;
            private bool _insertSpace;

            public void Append(string s)
            {
                if (string.IsNullOrEmpty(s)) return;
                if (_insertSpace)
                {
                    _stringBuilder.Append(' ');
                }
                else
                {
                    _insertSpace = true;
                }
                _stringBuilder.Append(s);
            }

            public override string ToString()
            {
                return _stringBuilder.ToString();
            }
        }

        internal static string ApplyCaps(StringBuilder sb, Заглавные заглавные)
        {
            заглавные.Применить(sb);
            return sb.ToString();
        }
    }

    public abstract class Заглавные
    {
        public abstract void Применить(StringBuilder sb);

        private class AllCharacters : Заглавные
        {
            public override void Применить(StringBuilder sb)
            {
                for (int i = 0; i < sb.Length; ++i)
                {
                    sb[i] = char.ToUpperInvariant(sb[i]);
                }
            }
        }

        private class NonCharacters : Заглавные
        {
            public override void Применить(StringBuilder sb)
            {
            }
        }

        private class OnlyFirstCharacter : Заглавные
        {
            public override void Применить(StringBuilder sb)
            {
                sb[0] = char.ToUpperInvariant(sb[0]);
            }
        }

        public static readonly Заглавные Все = new AllCharacters();
        public static readonly Заглавные Нет = new NonCharacters();
        public static readonly Заглавные Первая = new OnlyFirstCharacter();
    }

    public class Currency
    {
        public Currency(IЕдиницаИзмерения основная, IЕдиницаИзмерения дробная)
        {
            ОсновнаяЕдиница = основная;
            ДробнаяЕдиница = дробная;
        }

        public IЕдиницаИзмерения ОсновнаяЕдиница { get; }
        public IЕдиницаИзмерения ДробнаяЕдиница { get; }

        public static readonly Currency Рубли = new Currency(
            new ЕдиницаИзмерения(РодЧисло.Мужской, "рубль", "рубля", "рублей"),
            new ЕдиницаИзмерения(РодЧисло.Женский, "копейка", "копейки", "копеек"));

        public static readonly Currency Доллары = new Currency(
            new ЕдиницаИзмерения(РодЧисло.Мужской, "доллар США", "доллара США", "долларов США"),
            new ЕдиницаИзмерения(РодЧисло.Мужской, "цент", "цента", "центов"));

        public static readonly Currency Евро = new Currency(
            new ЕдиницаИзмерения(РодЧисло.Мужской, "евро", "евро", "евро"),
            new ЕдиницаИзмерения(РодЧисло.Мужской, "цент", "цента", "центов"));
    }

    public class ЕдиницаИзмерения : IЕдиницаИзмерения
    {
        public ЕдиницаИзмерения(
            РодЧисло родЧисло,
            string именЕдинственное,
            string родЕдинственное,
            string родМножественное)
        {
            _родЧисло = родЧисло;
            _именЕдинственное = именЕдинственное;
            _родЕдинственное = родЕдинственное;
            _родМножественное = родМножественное;
        }

        private readonly РодЧисло _родЧисло;
        private readonly string _именЕдинственное;
        private readonly string _родЕдинственное;
        private readonly string _родМножественное;
        string IЕдиницаИзмерения.ИменЕдинственное => _именЕдинственное;
        string IЕдиницаИзмерения.РодЕдинственное => _родЕдинственное;
        string IЕдиницаИзмерения.РодМножественное => _родМножественное;
        РодЧисло IЕдиницаИзмерения.РодЧисло => _родЧисло;
    }

    public abstract class РодЧисло : IЕдиницаИзмерения
    {
        internal abstract string ПолучитьФорму(IИзменяетсяПоРодам слово);

        private class Masculine : РодЧисло
        {
            internal override string ПолучитьФорму(IИзменяетсяПоРодам слово)
            {
                return слово.Мужской;
            }
        }

        private class Feminine : РодЧисло
        {
            internal override string ПолучитьФорму(IИзменяетсяПоРодам слово)
            {
                return слово.Женский;
            }
        }

        public static readonly РодЧисло Мужской = new Masculine();
        public static readonly РодЧисло Женский = new Feminine();

        РодЧисло IЕдиницаИзмерения.РодЧисло => this;
        string IЕдиницаИзмерения.ИменЕдинственное => null;
        string IЕдиницаИзмерения.РодЕдинственное => null;
        string IЕдиницаИзмерения.РодМножественное => null;
    }

    internal interface IИзменяетсяПоРодам
    {
        string Мужской { get; }
        string Женский { get; }
    }

    public interface IЕдиницаИзмерения
    {
        string ИменЕдинственное { get; }
        string РодЕдинственное { get; }
        string РодМножественное { get; }
        РодЧисло РодЧисло { get; }
    }
}
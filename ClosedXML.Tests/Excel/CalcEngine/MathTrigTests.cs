﻿// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class MathTrigTests
    {
        private const double tolerance = 1e-10;

        [Theory]
        public void Abs_ReturnsItselfOnPositiveNumbers([Range(0, 10, 0.1)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(input, actual, tolerance * 10);
        }

        [Theory]
        public void Abs_ReturnsTheCorrectValueOnNegativeInput([Range(-10, -0.1, 0.1)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(string.Format(@"ABS({0})", input.ToString(CultureInfo.InvariantCulture)));
            Assert.AreEqual(-input, actual, tolerance * 10);
        }

        [TestCase(-1, 3.141592654)]
        [TestCase(-0.9, 2.690565842)]
        [TestCase(-0.8, 2.498091545)]
        [TestCase(-0.7, 2.346193823)]
        [TestCase(-0.6, 2.214297436)]
        [TestCase(-0.5, 2.094395102)]
        [TestCase(-0.4, 1.982313173)]
        [TestCase(-0.3, 1.875488981)]
        [TestCase(-0.2, 1.772154248)]
        [TestCase(-0.1, 1.670963748)]
        [TestCase(0, 1.570796327)]
        [TestCase(0.1, 1.470628906)]
        [TestCase(0.2, 1.369438406)]
        [TestCase(0.3, 1.266103673)]
        [TestCase(0.4, 1.159279481)]
        [TestCase(0.5, 1.047197551)]
        [TestCase(0.6, 0.927295218)]
        [TestCase(0.7, 0.79539883)]
        [TestCase(0.8, 0.643501109)]
        [TestCase(0.9, 0.451026812)]
        [TestCase(1, 0)]
        public void Acos_ReturnsCorrectValue(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ACOS({input})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Acos_returns_error_when_number_outside_range([Range(1.1, 3, 0.1)] double input)
        {
            // checking input and it's additive inverse as both are outside range.
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOS({input})"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOS({-input})"));
        }

        [Theory]
        public void Acosh_NumbersBelow1ThrowNumberException([Range(-1, 0.9, 0.1)] double input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOSH({input})"));
        }

        [TestCase(1.2, 0.622362504)]
        [TestCase(1.5, 0.96242365)]
        [TestCase(1.8, 1.192910731)]
        [TestCase(2.1, 1.372859144)]
        [TestCase(2.4, 1.522079367)]
        [TestCase(2.7, 1.650193455)]
        [TestCase(3, 1.762747174)]
        [TestCase(3.3, 1.863279351)]
        [TestCase(3.6, 1.954207529)]
        [TestCase(3.9, 2.037266466)]
        [TestCase(4.2, 2.113748231)]
        [TestCase(4.5, 2.184643792)]
        [TestCase(4.8, 2.250731414)]
        [TestCase(5.1, 2.312634419)]
        [TestCase(5.4, 2.370860342)]
        [TestCase(5.7, 2.425828318)]
        [TestCase(6, 2.47788873)]
        [TestCase(1, 0)]
        public void Acosh_returns_correct_number(double angle, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ACOSH({angle})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [TestCase(-10, 3.041924001)]
        [TestCase(-9, 3.030935432)]
        [TestCase(-8, 3.017237659)]
        [TestCase(-7, 2.999695599)]
        [TestCase(-6, 2.976443976)]
        [TestCase(-5, 2.944197094)]
        [TestCase(-4, 2.89661399)]
        [TestCase(-3, 2.819842099)]
        [TestCase(-2, 2.677945045)]
        [TestCase(-1, 2.35619449)]
        [TestCase(0, 1.570796327)]
        [TestCase(1, 0.785398163)]
        [TestCase(2, 0.463647609)]
        [TestCase(3, 0.321750554)]
        [TestCase(4, 0.244978663)]
        [TestCase(5, 0.19739556)]
        [TestCase(6, 0.165148677)]
        [TestCase(7, 0.141897055)]
        [TestCase(8, 0.124354995)]
        [TestCase(9, 0.110657221)]
        [TestCase(10, 0.099668652)]
        public void Acot_returns_correct_number(double angle, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ACOT({angle})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Acoth_returns_error_for_absolute_angle_smaller_than_one([Range(-0.9, 0.9, 0.1)] double angle)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ACOTH({angle})"));
        }

        [TestCase(-10, -0.100335348)]
        [TestCase(-9, -0.111571776)]
        [TestCase(-8, -0.125657214)]
        [TestCase(-7, -0.143841036)]
        [TestCase(-6, -0.168236118)]
        [TestCase(-5, -0.202732554)]
        [TestCase(-4, -0.255412812)]
        [TestCase(-3, -0.34657359)]
        [TestCase(-2, -0.549306144)]
        [TestCase(2, 0.549306144)]
        [TestCase(3, 0.34657359)]
        [TestCase(4, 0.255412812)]
        [TestCase(5, 0.202732554)]
        [TestCase(6, 0.168236118)]
        [TestCase(7, 0.143841036)]
        [TestCase(8, 0.125657214)]
        [TestCase(9, 0.111571776)]
        [TestCase(10, 0.100335348)]
        [TestCase(1E+100, 1E-100)]
        public void Acoth_returns_correct_number(double angle, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ACOTH({angle})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [TestCase("LVII", 57)]
        [TestCase(@"mcmxii", 1912)]
        [TestCase("", 0)]
        [TestCase("-IV", -4)]
        [TestCase("   XIV   ", 14)]
        [TestCase(@"MCMLXXXIII ", 1983)]
        [TestCase(@"IIIIIIIIM", 992)]
        [TestCase(@"CIVIIX", 102)]
        [TestCase(@"IIX", 8)]
        [TestCase(@"VIII", 8)]
        public void Arabic_returns_correct_number(string roman, int arabic)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ARABIC(\"{roman}\")");
            Assert.AreEqual(arabic, actual);
        }

        [Test]
        public void Arabic_solitary_minus_is_not_valid_roman_number()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("ARABIC(\"-\")"));
        }

        [Test]
        public void Arabic_can_have_at_most_255_chars()
        {
            Assert.AreEqual(255000, XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 255)}\")"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 256)}\")"));
        }

        [TestCase("- I")]
        [TestCase("roman")]
        public void Arabic_returns_conversion_error_on_invalid_numbers(string invalidRoman)
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"ARABIC(\"{invalidRoman}\")"));
        }

        [TestCase(-1, -1.570796327)]
        [TestCase(-0.9, -1.119769515)]
        [TestCase(-0.8, -0.927295218)]
        [TestCase(-0.7, -0.775397497)]
        [TestCase(-0.6, -0.643501109)]
        [TestCase(-0.5, -0.523598776)]
        [TestCase(-0.4, -0.411516846)]
        [TestCase(-0.3, -0.304692654)]
        [TestCase(-0.2, -0.201357921)]
        [TestCase(-0.1, -0.100167421)]
        [TestCase(0, 0)]
        [TestCase(0.1, 0.100167421)]
        [TestCase(0.2, 0.201357921)]
        [TestCase(0.3, 0.304692654)]
        [TestCase(0.4, 0.411516846)]
        [TestCase(0.5, 0.523598776)]
        [TestCase(0.6, 0.643501109)]
        [TestCase(0.7, 0.775397497)]
        [TestCase(0.8, 0.927295218)]
        [TestCase(0.9, 1.119769515)]
        [TestCase(1, 1.570796327)]
        public void Asin_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ASIN({input})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Asin_ThrowsNumberExceptionWhenAbsOfInputGreaterThan1([Range(-3, -1.1, 0.1)] double input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ASIN({input})"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ASIN({-input})"));
        }

        [TestCase(0, 0)]
        [TestCase(0.1, 0.0998340788992076)]
        [TestCase(0.2, 0.198690110349241)]
        [TestCase(0.3, 0.295673047563422)]
        [TestCase(0.4, 0.390035319770715)]
        [TestCase(0.5, 0.481211825059603)]
        [TestCase(0.6, 0.568824898732248)]
        [TestCase(0.7, 0.652666566082356)]
        [TestCase(0.8, 0.732668256045411)]
        [TestCase(0.9, 0.808866935652783)]
        [TestCase(1, 0.881373587019543)]
        [TestCase(2, 1.44363547517881)]
        [TestCase(3, 1.81844645923207)]
        [TestCase(4, 2.0947125472611)]
        [TestCase(5, 2.31243834127275)]
        public void Asinh_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ASINH({input})");
            Assert.AreEqual(expectedResult, actual, tolerance);
            var minusActual = (double)XLWorkbook.EvaluateExpr($"ASINH({-input})");
            Assert.AreEqual(-expectedResult, minusActual, tolerance);
        }

        [TestCase(0, 0)]
        [TestCase(0.1, 0.099668652491162)]
        [TestCase(0.2, 0.197395559849881)]
        [TestCase(0.3, 0.291456794477867)]
        [TestCase(0.4, 0.380506377112365)]
        [TestCase(0.5, 0.463647609000806)]
        [TestCase(0.6, 0.540419500270584)]
        [TestCase(0.7, 0.610725964389209)]
        [TestCase(0.8, 0.674740942223553)]
        [TestCase(0.9, 0.732815101786507)]
        [TestCase(1, 0.785398163397448)]
        [TestCase(2, 1.10714871779409)]
        [TestCase(3, 1.24904577239825)]
        [TestCase(4, 1.32581766366803)]
        [TestCase(5, 1.37340076694502)]
        public void Atan_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN({input})");
            Assert.AreEqual(expectedResult, actual, tolerance);
            var minusActual = (double)XLWorkbook.EvaluateExpr($"ATAN({-input})");
            Assert.AreEqual(-expectedResult, minusActual, tolerance);
        }

        [Test]
        public void Atan2_Returns0OnSecond0AndFirstGreater0([Range(0.1, 5, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
            Assert.AreEqual(0, actual, tolerance);
        }

        [TestCase(1, 2, 1.10714871779409)]
        [TestCase(1, 3, 1.24904577239825)]
        [TestCase(2, 3, 0.98279372324733)]
        [TestCase(1, 4, 1.32581766366803)]
        [TestCase(3, 4, 0.92729521800161)]
        [TestCase(1, 5, 1.37340076694502)]
        [TestCase(2, 5, 1.19028994968253)]
        [TestCase(3, 5, 1.03037682652431)]
        [TestCase(4, 5, 0.89605538457134)]
        [TestCase(1, 6, 1.40564764938027)]
        [TestCase(5, 6, 0.87605805059819)]
        [TestCase(1, 7, 1.42889927219073)]
        [TestCase(2, 7, 1.29249666778979)]
        [TestCase(3, 7, 1.16590454050981)]
        [TestCase(4, 7, 1.05165021254837)]
        [TestCase(5, 7, 0.95054684081208)]
        [TestCase(6, 7, 0.86217005466723)]
        public void Atan2_ReturnsCorrectResults_EqualOnAllMultiplesOfFraction(double x, double y, double expectedResult)
        {
            for (int i = 1; i < 5; i++)
            {
                var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({x * i}, {y * i})");
                Assert.AreEqual(expectedResult, actual, tolerance);
            }
        }

        [Test]
        public void Atan2_ReturnsHalfPiOn0AsFirstInputWhenSecondGreater0([Range(0.1, 5, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
            Assert.AreEqual(0.5 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsMinus3QuartersOfPiWhenFirstSmaller0AndSecondItsNegative([Range(-5, -0.1, 0.3)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
            Assert.AreEqual(-0.75 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsMinusHalfPiOn0AsFirstInputWhenSecondSmaller0([Range(-5, -0.1, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
            Assert.AreEqual(-0.5 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsPiOn0AsSecondInputWhenFirstSmaller0([Range(-5, -0.1, 0.4)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
            Assert.AreEqual(Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ReturnsQuarterOfPiWhenInputsAreEqualAndGreater0([Range(0.1, 5, 0.3)] double input)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
            Assert.AreEqual(0.25 * Math.PI, actual, tolerance);
        }

        [Test]
        public void Atan2_ThrowsDiv0ExceptionOn0And0()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("ATAN2(0, 0)"));
        }

        [TestCase(-0.99, -2.64665241236225)]
        [TestCase(-0.9, -1.47221948958322)]
        [TestCase(-0.8, -1.09861228866811)]
        [TestCase(-0.6, -0.693147180559945)]
        [TestCase(-0.4, -0.423648930193602)]
        [TestCase(-0.2, -0.202732554054082)]
        [TestCase(0, 0)]
        [TestCase(0.2, 0.202732554054082)]
        [TestCase(0.4, 0.423648930193602)]
        [TestCase(0.6, 0.693147180559945)]
        [TestCase(0.8, 1.09861228866811)]
        [TestCase(-0.9, -1.47221948958322)]
        [TestCase(-0.990, -2.64665241236225)]
        [TestCase(-0.999, -3.8002011672502)]
        public void Atanh_ReturnsCorrectResults(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATANH({input})");
            Assert.AreEqual(expectedResult, actual, tolerance * 10);
        }

        [Theory]
        public void Atanh_ThrowsNumberExceptionWhenAbsOfInput1OrGreater([Range(1, 5, 0.2)] double input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ATANH({input})"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"ATANH({-input})"));
        }

        [TestCase(0, 36, "0")]
        [TestCase(1, 36, "1")]
        [TestCase(2, 36, "2")]
        [TestCase(3, 36, "3")]
        [TestCase(4, 36, "4")]
        [TestCase(5, 36, "5")]
        [TestCase(6, 36, "6")]
        [TestCase(7, 36, "7")]
        [TestCase(8, 36, "8")]
        [TestCase(9, 36, "9")]
        [TestCase(10, 36, "A")]
        [TestCase(11, 36, "B")]
        [TestCase(12, 36, "C")]
        [TestCase(13, 36, "D")]
        [TestCase(14, 36, "E")]
        [TestCase(15, 36, "F")]
        [TestCase(16, 36, "G")]
        [TestCase(17, 36, "H")]
        [TestCase(18, 36, "I")]
        [TestCase(19, 36, "J")]
        [TestCase(20, 36, "K")]
        [TestCase(21, 36, "L")]
        [TestCase(22, 36, "M")]
        [TestCase(23, 36, "N")]
        [TestCase(24, 36, "O")]
        [TestCase(25, 36, "P")]
        [TestCase(26, 36, "Q")]
        [TestCase(27, 36, "R")]
        [TestCase(28, 36, "S")]
        [TestCase(29, 36, "T")]
        [TestCase(30, 36, "U")]
        [TestCase(31, 36, "V")]
        [TestCase(32, 36, "W")]
        [TestCase(33, 36, "X")]
        [TestCase(34, 36, "Y")]
        [TestCase(35, 36, "Z")]
        [TestCase(36, 36, "10")]
        [TestCase(255, 29, "8N")]
        [TestCase(255, 2, "11111111")]
        public void Base_returns_number_in_specified_base(int input, int radix, string expectedResult)
        {
            var actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(255, 2, 3, "11111111")]
        [TestCase(255, 2, 8, "11111111")]
        [TestCase(255, 2, 10, "0011111111")]
        [TestCase(10, 3, 4, "0101")]
        [TestCase(0, 10, 0, "")]
        public void Base_returns_text_of_at_least_minimal_length(int input, int radix, int minLength, string expectedResult)
        {
            var actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})");
            Assert.AreEqual(expectedResult, actual);
        }

        [Test]
        public void Base_min_length_must_be_at_most_255()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("BASE(0,2,256)"));
        }

        [TestCase(@"""x""", "2", "2")]
        [TestCase("0", @"""x""", "2")]
        [TestCase("0", "2", @"""x""")]
        public void Base_coercion(string input, string radix, string minLength)
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})"));
        }

        [Theory]
        public void Base_radix_must_be_between_2_and_36([Range(-2, 1), Range(37, 40)] int radix)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"BASE(0,{radix})"));
        }

        [Theory]
        public void Base_number_must_be_zero_or_positive([Range(-5, -1)] int input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"BASE({input},2)"));
        }

        [Theory]
        public void Base_number_must_fit_in_double_without_precision_loss()
        {
            Assert.AreEqual(@"2GOPQOE5GCG", XLWorkbook.EvaluateExpr("BASE(9.007E+15,36)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("BASE(9.008E+15,36)"));
        }

        [TestCase(24.3, 5, 25)]
        [TestCase(6.7, 1, 7)]
        [TestCase(-8.1, 2, -8)]
        [TestCase(5.5, 2.1, 6.3)]
        [TestCase(5.5, 0, 0)]
        [TestCase(-5.5, 2.1, -4.2)]
        [TestCase(-5.5, -2.1, -6.3)]
        [TestCase(-5.5, 0, 0)]
        [TestCase(0, 0, 0)]
        [TestCase(0, 0.1, 0)]
        [TestCase(0, -0.1, 0)]
        [TestCase(0.1, 0, 0)]
        [TestCase(-0.1, 0, 0)]
        public void Ceiling(double input, double significance, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(6.7, -1)]
        [TestCase(0.1, -0.2)]
        public void Ceiling_returns_error_on_different_number_and_significance(double input, double significance)
        {
            // Spec says "if x and significance have different signs, #NUM! is returned.",
            // but in reality it only happens when number is positive and step negative.
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})"));
        }

        [TestCase(24.3, 5, null, 25)]
        [TestCase(6.7, null, null, 7)]
        [TestCase(-8.1, 2, null, -8)]
        [TestCase(-5.5, 2, -1, -6)]
        [TestCase(-5.5, 2, -0.1, -6)]
        [TestCase(5.5, 2.1, 0, 6.3)]
        [TestCase(5.5, -2.1, 0, 6.3)]
        [TestCase(5.5, 0, 0, 0)]
        [TestCase(5.5, 2.1, -1, 6.3)]
        [TestCase(5.5, -2.1, -1, 6.3)]
        [TestCase(5.5, 0, -1, 0)]
        [TestCase(5.5, 2.1, 10, 6.3)]
        [TestCase(5.5, -2.1, 10, 6.3)]
        [TestCase(5.5, 0, 10, 0)]
        [TestCase(-5.5, 2.1, 0, -4.2)]
        [TestCase(-5.5, -2.1, 0, -4.2)]
        [TestCase(-5.5, 0, 0, 0)]
        [TestCase(-5.5, 2.1, -1, -6.3)]
        [TestCase(-5.5, -2.1, -1, -6.3)]
        [TestCase(-5.5, 0, -1, 0)]
        [TestCase(-5.5, 2.1, 10, -6.3)]
        [TestCase(-5.5, -2.1, 10, -6.3)]
        [TestCase(-5.5, 0, 10, 0)]
        public void CeilingMath(double input, double? significance, double? mode, double expectedResult)
        {
            var parameters = new StringBuilder();
            parameters.Append(input);
            if (significance != null)
            {
                parameters.Append(", ").Append(significance);
                if (mode != null)
                    parameters.Append(", ").Append(mode);
            }

            var actual = (double)XLWorkbook.EvaluateExpr($"CEILING.MATH({parameters})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [Test]
        public void Combin()
        {
            var actual1 = XLWorkbook.EvaluateExpr("COMBIN(200, 2)");
            Assert.AreEqual(19900.0, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("COMBIN(20.1, 2.9)");
            Assert.AreEqual(190.0, actual2);
        }

        [Theory]
        public void Combin_returns_1_for_k_is_0_or_k_equals_n([Range(0, 10)] int n)
        {
            var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 0)");
            Assert.AreEqual(1, actual);

            var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n})");
            Assert.AreEqual(1, actual2);
        }

        [TestCase(0, 0, 1)]
        [TestCase(1, 0, 1)]
        [TestCase(1, 1, 1)]
        [TestCase(4, 2, 6)]
        [TestCase(5, 2, 10)]
        [TestCase(6, 2, 15)]
        [TestCase(6, 3, 20)]
        [TestCase(7, 2, 21)]
        [TestCase(7, 3, 35)]
        public void Combin_calculates_combinations(int n, int k, int expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, {k})");
            Assert.AreEqual(expectedResult, actual);

            var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - k})");
            Assert.AreEqual(expectedResult, actual2);
        }

        [Theory]
        public void Combin_returns_n_for_k_is_1_or_k_is_n_minus_1([Range(1, 10)] int n)
        {
            var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 1)");
            Assert.AreEqual(n, actual);

            var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - 1})");
            Assert.AreEqual(n, actual2);
        }

        [Test]
        public void Combin_returns_num_error_when_k_is_larger_than_n()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(5, 6)"));

            // Values are floored, so this is COMBIN(5, 5).
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("COMBIN(5, 5.5)"));
        }

        [Test]
        public void Combin_returns_num_error_when_value_is_too_large()
        {
            // Maximum int - 1 is maximum computable value in Excel.
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(2147483647, 2147483647)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(5E+301, 6)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("COMBIN(6, 5E+301)"));
        }

        [TestCase(-4)]
        [TestCase(-3)]
        [TestCase(-1)]
        [TestCase(-0.1)]
        public void Combin_returns_num_error_for_any_argument_smaller_than_0(double smaller0)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    smaller0.ToString(CultureInfo.InvariantCulture),
                    (-smaller0).ToString(CultureInfo.InvariantCulture))));

            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr(
                string.Format(
                    @"COMBIN({0}, {1})",
                    (-smaller0).ToString(CultureInfo.InvariantCulture),
                    smaller0.ToString(CultureInfo.InvariantCulture))));
        }

        [TestCase("\"no number\"")]
        [TestCase("\"\"")]
        public void Combin_returns_value_error_for_any_non_numeric_argument(string input)
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"COMBIN({input}, 1)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"COMBIN(1, {input})"));
        }

        [TestCase(4, 3, 20)]
        [TestCase(10, 3, 220)]
        [TestCase(0, 0, 1)]
        [TestCase(1, 0, 1)]
        [TestCase(10, 15, 1307504)]
        public void Combina_calculates_correct_values(int number, int chosen, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Theory]
        public void Combina_returns_one_when_chosen_is_zero([Range(0, 10)] int number)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, 0)");
            Assert.AreEqual(1, actualResult);
        }

        [TestCase(-1, 2)]
        [TestCase(-3, -2)]
        [TestCase(2, -2)]
        [TestCase(int.MaxValue + 1d, 1)]
        public void Combina_returns_error_on_invalid_values(double number, int chosen)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})"));
        }

        [TestCase(4.23, 3, 20)]
        [TestCase(10.4, 3.14, 220)]
        [TestCase(0, 0.4, 1)]
        public void Combina_truncates_numbers_to_zero(double number, double chosen, int expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [TestCase(0, 1)]
        [TestCase(0.4, 0.921060994002885)]
        [TestCase(0.8, 0.696706709347165)]
        [TestCase(1.2, 0.362357754476674)]
        [TestCase(1.6, -0.0291995223012888)]
        [TestCase(2, -0.416146836547142)]
        [TestCase(2.4, -0.737393715541245)]
        [TestCase(2.8, -0.942222340668658)]
        [TestCase(3.2, -0.998294775794753)]
        [TestCase(3.6, -0.896758416334147)]
        [TestCase(4, -0.653643620863612)]
        [TestCase(4.4, -0.307332869978419)]
        [TestCase(4.8, 0.0874989834394464)]
        [TestCase(5.2, 0.468516671300377)]
        [TestCase(5.6, 0.77556587851025)]
        [TestCase(6, 0.960170286650366)]
        [TestCase(6.4, 0.993184918758193)]
        [TestCase(6.8, 0.869397490349825)]
        [TestCase(7.2, 0.608351314532255)]
        [TestCase(7.6, 0.251259842582256)]
        [TestCase(8, -0.145500033808614)]
        [TestCase(8.4, -0.519288654116686)]
        public void Cos_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actualResult = (double)XLWorkbook.EvaluateExpr($"COS({input})");
            Assert.AreEqual(expectedResult, actualResult, tolerance);
        }

        [TestCase(0, 1)]
        [TestCase(0.4, 1.08107237183845)]
        [TestCase(0.8, 1.33743494630484)]
        [TestCase(1.2, 1.81065556732437)]
        [TestCase(1.6, 2.57746447119489)]
        [TestCase(2, 3.76219569108363)]
        [TestCase(2.4, 5.55694716696551)]
        [TestCase(2.8, 8.25272841686113)]
        [TestCase(3.2, 12.2866462005439)]
        [TestCase(3.6, 18.3127790830626)]
        [TestCase(4, 27.3082328360165)]
        [TestCase(4.4, 40.7315730024356)]
        [TestCase(4.8, 60.7593236328919)]
        [TestCase(5.2, 90.638879219786)]
        [TestCase(5.6, 135.215052644935)]
        [TestCase(6, 201.715636122456)]
        [TestCase(6.4, 300.923349714678)]
        [TestCase(6.8, 448.924202712783)]
        [TestCase(7.2, 669.715755490113)]
        [TestCase(7.6, 999.098197777775)]
        [TestCase(8, 1490.47916125218)]
        [TestCase(8.4, 2223.53348628359)]
        public void Cosh_ReturnsCorrectResult(double input, double expectedResult)
        {
            var actualResult = (double)XLWorkbook.EvaluateExpr($"COSH({input})");
            Assert.AreEqual(expectedResult, actualResult, tolerance);
            var actualResult2 = (double)XLWorkbook.EvaluateExpr($"COSH({-input})");
            Assert.AreEqual(expectedResult, actualResult2, tolerance);
        }

        [TestCase(711)]
        [TestCase(-711)]
        [TestCase(100000)]
        public void Cosh_too_large_returns_error(double input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"COSH({input})"));
        }

        [TestCase(1, 0.642092616)]
        [TestCase(2, -0.457657554)]
        [TestCase(3, -7.015252551)]
        [TestCase(4, 0.863691154)]
        [TestCase(5, -0.295812916)]
        [TestCase(6, -3.436353004)]
        [TestCase(7, 1.147515422)]
        [TestCase(8, -0.147065064)]
        [TestCase(9, -2.210845411)]
        [TestCase(10, 1.542351045)]
        [TestCase(11, -0.004425741)]
        [TestCase(Math.PI * 0.5, 0)]
        [TestCase(45, 0.617369624)]
        [TestCase(-2, 0.457657554)]
        [TestCase(-3, 7.015252551)]
        public void Cot(double angle, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"COT({angle})");
            Assert.AreEqual(expected, actual, tolerance * 10.0);
        }

        [Test]
        public void Cot_returns_division_by_zero_error_on_angle_zero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COT(0)"));
        }

        [Test]
        public void Coth_returns_division_by_zero_error_on_angle_zero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COTH(0)"));
        }

        [TestCase(-10, -1.000000004)]
        [TestCase(-9, -1.00000003)]
        [TestCase(-8, -1.000000225)]
        [TestCase(-7, -1.000001663)]
        [TestCase(-6, -1.000012289)]
        [TestCase(-5, -1.000090804)]
        [TestCase(-4, -1.00067115)]
        [TestCase(-3, -1.004969823)]
        [TestCase(-2, -1.037314721)]
        [TestCase(-1, -1.313035285)]
        [TestCase(1, 1.313035285)]
        [TestCase(2, 1.037314721)]
        [TestCase(3, 1.004969823)]
        [TestCase(4, 1.00067115)]
        [TestCase(5, 1.000090804)]
        [TestCase(6, 1.000012289)]
        [TestCase(7, 1.000001663)]
        [TestCase(8, 1.000000225)]
        [TestCase(9, 1.00000003)]
        [TestCase(10, 1.000000004)]
        public void Coth_returns_correct_number(double angle, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"COTH({angle})");
            Assert.AreEqual(expected, actual, tolerance * 10.0);
        }

        [Test]
        public void Csc_returns_division_by_zero_on_angle_zero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("CSC(0)"));
        }

        [TestCase(-10, 1.838163961)]
        [TestCase(-9, -2.426486644)]
        [TestCase(-8, -1.010756218)]
        [TestCase(-7, -1.522101063)]
        [TestCase(-6, 3.578899547)]
        [TestCase(-5, 1.042835213)]
        [TestCase(-4, 1.321348709)]
        [TestCase(-3, -7.086167396)]
        [TestCase(-2, -1.09975017)]
        [TestCase(-1, -1.188395106)]
        [TestCase(1, 1.188395106)]
        [TestCase(2, 1.09975017)]
        [TestCase(3, 7.086167396)]
        [TestCase(4, -1.321348709)]
        [TestCase(5, -1.042835213)]
        [TestCase(6, -3.578899547)]
        [TestCase(7, 1.522101063)]
        [TestCase(8, 1.010756218)]
        [TestCase(9, 2.426486644)]
        [TestCase(10, -1.838163961)]
        public void Csc_returns_correct_number(double angle, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"CSC({angle})");
            Assert.AreEqual(expected, actual, tolerance * 10);
        }

        [TestCase(1, 0.850918128)]
        [TestCase(2, 0.275720565)]
        [TestCase(3, 0.09982157)]
        [TestCase(4, 0.03664357)]
        [TestCase(5, 0.013476506)]
        [TestCase(6, 0.004957535)]
        [TestCase(7, 0.001823765)]
        [TestCase(8, 0.000670925)]
        [TestCase(9, 0.00024682)]
        [TestCase(10, 0.000090799859712122200000)]
        [TestCase(11, 0.0000334034)]
        public void Csch_calculates_correct_values(double input, double expectedOutput)
        {
            Assert.AreEqual(expectedOutput, (double)XLWorkbook.EvaluateExpr($"CSCH({input})"), 0.000000001);
        }

        [Test]
        public void Csch_returns_division_error_on_angle_zero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("CSCH(0)"));
        }

        [TestCase("FF", 16, 255)]
        [TestCase("111", 2, 7)]
        [TestCase("zap", 36, 45745)] // Case insensitive
        [TestCase("  1234", 10, 1234)] // Trims start
        [TestCase("123", 10.9, 123)] // Radix truncated
        [TestCase("1F", 10, XLError.NumberInvalid)]
        [TestCase("", 10, 0)]
        public void Decimal(string inputString, double radix, object expectedResult)
        {
            var actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Theory]
        public void Decimal_radix_must_be_between_2_and_36([Range(37, 255), Range(-5, 1)] int radix)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Test]
        public void Decimal_zero_is_zero_in_any_radix([Range(2, 36)] int radix)
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})"));
        }

        [Test]
        public void Decimal_text_must_be_less_than_256_chars_long()
        {
            var text = new string('0', 256);
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"DECIMAL(\"{text}\", 10)"));
        }

        [Test]
        public void Decimal_returns_number_invalid_when_result_out_of_bounds()
        {
            Assert.AreEqual(1.4057081148316923E+308d, (double)XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 198)}\", 36)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 199)}\", 36)"));
        }

        [TestCase("101", "\"1 2/2\"", 5)] // 101 in binary is 5
        public void Decimal_coercion(string input, string radix, object expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr($"DECIMAL({input}, {radix})"));
        }

        [Test]
        public void Degrees()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("DEGREES(PI())");
            Assert.AreEqual(180, actual, XLHelper.Epsilon);
        }

        [TestCase(0, 0)]
        [TestCase(Math.PI, 180)]
        [TestCase(Math.PI * 2, 360)]
        [TestCase(1, 57.2957795130823)]
        [TestCase(2, 114.591559026165)]
        [TestCase(3, 171.887338539247)]
        [TestCase(4, 229.183118052329)]
        [TestCase(5, 286.478897565412)]
        [TestCase(6, 343.774677078494)]
        [TestCase(7, 401.070456591576)]
        [TestCase(8, 458.366236104659)]
        [TestCase(9, 515.662015617741)]
        [TestCase(10, 572.957795130823)]
        [TestCase(Math.PI * 0.5, 90)]
        [TestCase(Math.PI * 1.5, 270)]
        [TestCase(Math.PI * 0.25, 45)]
        [TestCase(-1, -57.2957795130823)]
        public void Degrees_ReturnsCorrectResult(double input, double expected)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"DEGREES({input})");
            Assert.AreEqual(expected, actual, tolerance);
        }

        [TestCase(3, 4)]
        [TestCase(2, 2)]
        [TestCase(-1, -2)]
        [TestCase(-2, -2)]
        [TestCase(0, 0)]
        [TestCase(1.5, 2)]
        [TestCase(2.01, 4)]
        [TestCase(1e+100, 1e+100)]
        [TestCase(Math.PI, 4)]
        public void Even(double number, double expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr($"EVEN({number})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(0, 1)]
        [TestCase(1, Math.E)]
        [TestCase(2, 7.38905609893065)]
        [TestCase(3, 20.0855369231877)]
        [TestCase(4, 54.5981500331442)]
        [TestCase(5, 148.413159102577)]
        [TestCase(6, 403.428793492735)]
        [TestCase(7, 1096.63315842846)]
        [TestCase(8, 2980.95798704173)]
        [TestCase(9, 8103.08392757538)]
        [TestCase(10, 22026.4657948067)]
        [TestCase(11, 59874.1417151978)]
        [TestCase(12, 162754.791419004)]
        [TestCase(-1E+100, 0)]
        public void Exp_returns_correct_results(double input, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"EXP({input})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(710)]
        public void Exp_with_too_large_result_return_error(double input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"EXP({input})"));
        }

        [Test]
        public void Fact()
        {
            object actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
            Assert.AreEqual(120.0, actual);
        }

        [TestCase(0, 1d)]
        [TestCase(1, 1d)]
        [TestCase(2, 2d)]
        [TestCase(3, 6d)]
        [TestCase(4, 24d)]
        [TestCase(5, 120d)]
        [TestCase(6, 720d)]
        [TestCase(7, 5040d)]
        [TestCase(8, 40320d)]
        [TestCase(9, 362880d)]
        [TestCase(10, 3628800d)]
        [TestCase(11, 39916800d)]
        [TestCase(12, 479001600d)]
        [TestCase(13, 6227020800d)]
        [TestCase(14, 87178291200d)]
        [TestCase(15, 1307674368000d)]
        [TestCase(16, 20922789888000d)]
        [TestCase(170.9, 7.257415615308004E+306)]
        [TestCase(0.1, 1L)]
        [TestCase(2.3, 2L)]
        [TestCase(2.8, 2L)]
        public void Fact_calculates_factorial(double input, double expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr($@"FACT({input.ToString(CultureInfo.InvariantCulture)})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(-10)]
        [TestCase(-5)]
        [TestCase(-1)]
        [TestCase(-0.1)]
        public void Fact_returns_error_for_negative_input(double input)
        {
            var actual = XLWorkbook.EvaluateExpr($@"FACT({input.ToString(CultureInfo.InvariantCulture)})");
            Assert.AreEqual(XLError.NumberInvalid, actual);
        }

        [TestCase(171)]
        [TestCase(5000)]
        public void Fact_returns_error_for_too_large_result(int input)
        {
            var actual = XLWorkbook.EvaluateExpr($@"FACT({input})");
            Assert.AreEqual(XLError.NumberInvalid, actual);
        }

        [Test]
        public void Fact_coercion_fails_for_non_numeric_input()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"FACT(""x"")"));
        }

        [TestCase(0, 1L)]
        [TestCase(1, 1L)]
        [TestCase(2, 2L)]
        [TestCase(3, 3L)]
        [TestCase(4, 8L)]
        [TestCase(5, 15L)]
        [TestCase(6, 48L)]
        [TestCase(7, 105L)]
        [TestCase(8, 384L)]
        [TestCase(9, 945L)]
        [TestCase(10, 3840L)]
        [TestCase(11, 10395L)]
        [TestCase(12, 46080L)]
        [TestCase(13, 135135L)]
        [TestCase(14, 645120)]
        [TestCase(15, 2027025)]
        [TestCase(16, 10321920)]
        [TestCase(-1, 1L)]
        [TestCase(0, 1)]
        [TestCase(0.1, 1L)]
        [TestCase(1.4, 1L)]
        [TestCase(2.3, 2L)]
        [TestCase(2.8, 2L)]
        public void FactDouble_ReturnsCorrectResult(double input, long expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase(301)]
        [TestCase(1e+100)]
        public void FactDouble_returns_error_on_too_large_value(double n)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"FACTDOUBLE({n})"));
        }

        [Theory]
        public void FactDouble_ThrowsNumberExceptionForInputSmallerThanMinus1([Range(-10, -2)] int input)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})"));
        }

        [Test]
        public void FactDouble_ThrowsValueExceptionForNonNumericInput()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"FACTDOUBLE(""x"")"));
        }

        [TestCase(0, 0, 0)]
        [TestCase(0, 1, 0)]
        [TestCase(24.3, 5, 20)]
        [TestCase(6.7, 1, 6)]
        [TestCase(-8.1, 2, -10)]
        [TestCase(5.5, 2.1, 4.2)]
        [TestCase(-5.5, 2.1, -6.3)]
        [TestCase(-5.5, -2.1, -4.2)]
        public void Floor(double input, double significance, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})");
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [TestCase(6.7, 0)]
        [TestCase(-6.7, 0)]
        public void Floor_ThrowsDivisionByZeroOnZeroSignificance(double input, double significance)
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})"));
        }

        [TestCase(6.7, -1)]
        public void Floor_ThrowsNumberExceptionOnInvalidInput(double input, double significance)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})"));
        }

        [Test]
        // Functions have to support a period first before we can implement this
        [TestCase(24.3, 5, null, 20)]
        [TestCase(6.7, null, null, 6)]
        [TestCase(-8.1, 2, null, -10)]
        [TestCase(5.5, 2.1, 0, 4.2)]
        [TestCase(5.5, -2.1, 0, 4.2)]
        [TestCase(5.5, 0, 0, 0)]
        [TestCase(5.5, 2.1, -1, 4.2)]
        [TestCase(5.5, -2.1, -1, 4.2)]
        [TestCase(5.5, 0, -2, 0)]
        [TestCase(5.5, 2.1, 10, 4.2)]
        [TestCase(5.5, -2.1, 10, 4.2)]
        [TestCase(5.5, 0, 10, 0)]
        [TestCase(-5.5, 2.1, 0, -6.3)]
        [TestCase(-5.5, -2.1, 0, -6.3)]
        [TestCase(-5.5, 0, 0, 0)]
        [TestCase(-5.5, 2.1, -1, -4.2)]
        [TestCase(-5.5, -2.1, -1, -4.2)]
        [TestCase(-5.5, 0, -1, 0)]
        [TestCase(-5.5, 2.1, 10, -4.2)]
        [TestCase(-5.5, -2.1, 10, -4.2)]
        [TestCase(-5.5, 0, 0, 0)]
        [DefaultFloatingPointTolerance(tolerance)]
        public void FloorMath(double input, double? significance, int? mode, double expectedResult)
        {
            var parameters = new StringBuilder();
            parameters.Append(input);
            if (significance != null)
            {
                parameters.Append(", ").Append(significance);
                if (mode != null)
                    parameters.Append(", ").Append(mode);
            }

            var actual = (double)XLWorkbook.EvaluateExpr($"FLOOR.MATH({parameters})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase("24,36", ExpectedResult = 12)]
        [TestCase("240,360,30", ExpectedResult = 30)]
        [TestCase("24.9,36.9", ExpectedResult = 12)]
        [TestCase("{24,36}", ExpectedResult = 12)]
        [TestCase("{\"24\",\"36\"}", ExpectedResult = 12)]
        [TestCase("5,0", ExpectedResult = 5)]
        [TestCase("0,5", ExpectedResult = 5)]
        public double Gcd(string args)
        {
            return (double)XLWorkbook.EvaluateExpr($"GCD({args})");
        }

        [Test]
        public void Gcd_accepts_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                (120, 240),
                ("60", "150"),
            });
            Assert.AreEqual(30, ws.Evaluate("GCD(A1:A2,B1:B2)"));

            // Blank is considered 0
            Assert.AreEqual(60, ws.Evaluate("GCD(A1:A3)"));

            // Logical are not converted
            ws.Cell("A3").Value = true;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("GCD(A1:A3)"));

            // Unconvertable text causes error
            ws.Cell("A3").Value = "one";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("GCD(A1:A3)"));
        }

        [TestCase]
        public void Gcd_numbers_must_fit_in_double_without_precision_loss()
        {
            Assert.AreEqual(9.007E+15, XLWorkbook.EvaluateExpr("GCD(9.007E+15)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("GCD(9.008E+15)"));
        }

        [TestCase]
        public void Gcd_numbers_must_be_zero_or_positive()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("GCD(-1)"));
        }

        [TestCase(8.9, 8)]
        [TestCase(-8.9, -9)]
        public void Int(double input, double expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"INT({input})");
            Assert.AreEqual(expected, actual);
        }

        [TestCase("24, 36", ExpectedResult = 72)]
        [TestCase("24.9, 36.9", ExpectedResult = 72)]
        [TestCase("{24, 36}", ExpectedResult = 72)]
        [TestCase("{1,2,3;4,5,6}", ExpectedResult = 60)]
        [TestCase("{\"1\",\"2\",\"3\"}", ExpectedResult = 6)]
        [TestCase("240, 360, 30", ExpectedResult = 720)]
        [TestCase("5, 0", ExpectedResult = 0)]
        [TestCase("0, 5", ExpectedResult = 0)]
        public double Lcm(string args)
        {
            return (double)XLWorkbook.EvaluateExpr($"LCM({args})");
        }

        [Test]
        public void Lcm_accepts_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                (1, 2, 3),
                ("4", "5", "6"),
            });
            Assert.AreEqual(60, ws.Evaluate("LCM(A1:B2,C1:C2)"));

            // Blank is considered 0
            Assert.AreEqual(0, ws.Evaluate("LCM(A1:A3)"));

            // Logical are not converted
            ws.Cell("A3").Value = true;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("LCM(A1:A3)"));

            // Unconvertable text causes error
            ws.Cell("A3").Value = "one";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("LCM(A1:A3)"));
        }

        [TestCase]
        public void Lcm_numbers_must_fit_in_double_without_precision_loss()
        {
            Assert.AreEqual(9.007E+15, XLWorkbook.EvaluateExpr("LCM(9.007E+15)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LCM(9.008E+15)"));
        }

        [TestCase]
        public void Lcm_numbers_must_be_zero_or_positive()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LCM(-1)"));
        }

        [TestCase(86, 4.4543472962)]
        [TestCase(2.7182818, 0.9999999895)]
        [TestCase(20.085536923, 3)]
        public void Ln_calculates_logarithm(double x, double ln)
        {
            Assert.AreEqual(ln, (double)XLWorkbook.EvaluateExpr($"LN({x})"), tolerance);
        }

        [TestCase(0)]
        [TestCase(-0.7)]
        [TestCase(-10)]
        public void Ln_non_positive_returns_error(double x)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"LN({x})"));
        }

        [TestCase(10, 10, 1)]
        [TestCase(8, 2, 3)]
        [TestCase(86, 2.7182818, 4.4543473428883)]
        public void Log_calculates_logarithm(double x, double @base, double result)
        {
            Assert.AreEqual(result, (double)XLWorkbook.EvaluateExpr($"LOG({x}, {@base})"), tolerance);
        }

        [Test]
        public void Log_default_base_is_10()
        {
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("LOG(100)"));
        }

        [Test]
        public void Log_error_conditions()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(1,0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("LOG(0,0)"));
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("LOG(10,1)"));
        }

        [TestCase(86, 1.93449845124)]
        [TestCase(10, 1)]
        [TestCase(1E5, 5)]
        public void Log10_calculates_logarithm(double x, double expectedResult)
        {
            Assert.AreEqual(expectedResult, (double)XLWorkbook.EvaluateExpr($"LOG10({x})"), tolerance);
        }

        [TestCase(0)]
        [TestCase(-5)]
        [TestCase(-0.5)]
        public void Log10_error_conditions(double x)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"LOG10({x})"));
        }

        [Test]
        public void Log10_is_detected_inside_expression()
        {
            // Because LOG10 is extracted from CellFunction, make sure it is properly read even in the middle of expression.
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("0 + LOG10(10)"));
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void MDeterm()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                (2, 4),
                (3, 5),
            });

            ws.Cell("A5").FormulaA1 = "MDETERM(A1:B2)";
            var actual = ws.Cell("A5").Value;
            Assert.AreEqual(-2, (double)actual);

            ws.Cell("A6").FormulaA1 = "SUM(A5)";
            actual = ws.Cell("A6").Value;
            Assert.AreEqual(-2, (double)actual);

            ws.Cell("A7").FormulaA1 = "SUM(MDETERM(A1:B2))";
            actual = ws.Cell("A7").Value;
            Assert.AreEqual(-2, (double)actual);
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void MDeterm_examples()
        {
            // Examples from spec
            Assert.AreEqual(1, (double)XLWorkbook.EvaluateExpr("MDETERM({3,6,1;1,1,0;3,10,2})"));
            Assert.AreEqual(-3, XLWorkbook.EvaluateExpr("MDETERM({3,6;1,1})"));

            // Example from office website
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                ("Data", "Data", "Data", "Data"),
                (1, 3, 8, 5),
                (1, 3, 6, 1),
                (1, 1, 1, 0),
                (7, 3, 10, 2),
            });
            Assert.AreEqual(88, (double)ws.Evaluate("MDETERM(A2:D5)"));
        }

        [Test]
        public void MDeterm_requires_equal_number_of_rows_and_columns()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("MDETERM({1,2})"));
        }

        [Test]
        public void MDeterm_singular_matrix_returns_zero()
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr("MDETERM({1,2;1,2})"));
        }

        [Test]
        public void MDeterm_requires_all_array_elements_are_numbers()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                (2, 4),
                (3, 5),
            });

            ws.Cell("B2").Value = Blank.Value;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

            ws.Cell("B2").Value = "1";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

            ws.Cell("B2").Value = true;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MDETERM(A1:B2)"));

            ws.Cell("B2").Value = XLError.NameNotRecognized;
            Assert.AreEqual(XLError.NameNotRecognized, ws.Evaluate("MDETERM(A1:B2)"));
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void MInverse()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[]
            {
                (1, 2, 1),
                (3, 4, -1),
                (0, 2, 0),
            });

            ws.Cell("A5").FormulaA1 = "MINVERSE(A1:C3)";
            var actual = ws.Cell("A5").Value;
            Assert.AreEqual(0.25, (double)actual);

            ws.Cell("A6").FormulaA1 = "SUM(A5)";
            actual = ws.Cell("A6").Value;
            Assert.AreEqual(0.25, (double)actual);

            ws.Cell("A7").FormulaA1 = "SUM(MINVERSE(A1:C3))";
            actual = ws.Cell("A7").Value;
            Assert.AreEqual(0.5, (double)actual);
        }

        [Test]
        public void MInverse_returns_error_on_singular_matrix()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[]
            {
                (1, 2),
                (1, 2),
            });
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("MINVERSE(A1:B2)"));
        }

        [Test]
        public void MInverse_requires_square_matrix()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("MINVERSE({1,2,3;7,5,5})"));
        }

        [Test]
        public void MInverse_all_array_elements_must_be_numbers()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[]
            {
                (1, 2),
                (8, 4),
            });

            ws.Cell("B2").Value = Blank.Value;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

            ws.Cell("B2").Value = true;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

            ws.Cell("B2").Value = "1";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MINVERSE(A1:B2)"));

            ws.Cell("B2").Value = XLError.DivisionByZero;
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("MINVERSE(A1:B2)"));
        }

        [Test]
        public void MMult()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[]
            {
                (2, 4),
                (3, 5),
                (2, 4),
                (3, 5),
            });

            ws.Cell("A5").FormulaA1 = "MMULT(A1:B2, A3:B4)";
            var actual = ws.Cell("A5").Value;
            Assert.AreEqual(16.0, actual);

            ws.Cell("A6").FormulaA1 = "SUM(A5)";
            actual = ws.Cell("A6").Value;
            Assert.AreEqual(16.0, actual);

            ws.Cell("A7").FormulaA1 = "SUM(MMULT(A1:B2, A3:B4))";
            actual = ws.Cell("A7").Value;
            Assert.AreEqual(102.0, actual);
        }

        [Test]
        public void MMult_handles_non_square_matrices()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[]
            {
                // 2x3
                (1, 3, 5),
                (2, 4, 6),
                // 3x4
                (10, 13, 16, 19),
                (11, 14, 17, 20),
                (12, 15, 18, 21),
            });

            // 2x4 output expected:
            // 103, 130, 157, 184
            // 136, 172, 208, 244
            ws.Cell("A6").FormulaA1 = "MMult(A1:C2, A3:D5)";
            var actual = ws.Cell("A6").Value;
            Assert.AreEqual(103.0, actual);

            ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:C2, A3:D5))";
            actual = ws.Cell("A7").Value;
            Assert.AreEqual(1334, actual);
        }

        [TestCase("A2:C2", "A3:C3")] // 1x3 and 1x3
        [TestCase("A2:C4", "A5:C5")] // 3x3 and 1x3
        [TestCase("A2:C5", "A6:D6")] // 3x4 and 1x4
        public void MMult_array1_rows_must_match_array2_column(string array1Range, string array2Range)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells($"{array1Range}").Value = 1.0;
            ws.Cells($"{array2Range}").Value = 1.0;

            ws.Cell("A1").FormulaA1 = $"MMULT({array1Range},{array2Range})";

            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A1").Value);
        }

        [TestCase("")]
        [TestCase("Text")]
        public void MMult_ThrowsWhenCellsContainInvalidInput(string invalidInput)
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");

            // 2x3
            ws.Cell("A1").SetValue(1).CellRight().SetValue(3).CellRight().SetValue(invalidInput);
            ws.Cell("A2").SetValue(2).CellRight().SetValue(4).CellRight().SetValue(6);

            // 3x4
            ws.Cell("A3").SetValue(10).CellRight().SetValue(13).CellRight().SetValue(16).CellRight().SetValue(19);
            ws.Cell("A4").SetValue(11).CellRight().SetValue(14).CellRight().SetValue(17).CellRight().SetValue(20);
            ws.Cell("A5").SetValue(12).CellRight().SetValue(15).CellRight().SetValue(18).CellRight().SetValue(21);

            ws.Cell("A6").FormulaA1 = "MMULT(A1:C2,A3:D4)";

            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A6").Value);
        }

        [TestCase(1.5, 1, 0.5)]
        [TestCase(3, 2, 1)]
        [TestCase(-3, 2, 1)]
        [TestCase(-3, -2, -1)]
        [TestCase(-4.3, -0.5, -0.3)]
        [TestCase(6.9, -0.2, -0.1)]
        [TestCase(0.7, 0.6, 0.1)]
        [TestCase(6.2, 1.1, 0.7)]
        public void Mod(double x, double y, double result)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"MOD({x}, {y})");
            Assert.AreEqual(result, actual, tolerance);
        }

        [Test]
        public void Mod_divisor_zero_returns_error()
        {
            // Spec says that "If y is 0, the return value is unspecified", but Excel says #DIV/0!, so let's go with that.
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("MOD(1, 0)"));
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("MOD(0, 0)"));
        }

        [TestCase(10, 3, ExpectedResult = 9.0)]
        [TestCase(10.5, 3, ExpectedResult = 12.0)]
        [TestCase(10.4, 3, ExpectedResult = 9.0)]
        [TestCase(-10, -3, ExpectedResult = -9.0)]
        [TestCase(1.3, 0.2, ExpectedResult = 1.4)]
        [TestCase(5677.912288, 10, ExpectedResult = 5680.0)]
        [TestCase(5674.912288, 10, ExpectedResult = 5670.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(0.5, 1, ExpectedResult = 1.0)]
        [TestCase(0.49999, 1, ExpectedResult = 0.0)]
        [TestCase(-13.4, -3, ExpectedResult = -12.0)]
        [TestCase(-13.5, -3, ExpectedResult = -15.0)]
        [TestCase(0.9, 0.2, ExpectedResult = 1.0)]
        [TestCase(0.89999, 0.2, ExpectedResult = 0.8)]
        [TestCase(15.5, 3, ExpectedResult = 15.0)]
        [TestCase(1.4, 0.5, ExpectedResult = 1.5)]
        [TestCase(3, 7, ExpectedResult = 0)]
        [TestCase(3, 0, ExpectedResult = 0)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double MRound(double number, double multiple)
        {
            return (double)XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})");
        }

        [TestCase(123456.123, -10)]
        [TestCase(-123456.123, 5)]
        public void MRoundExceptions(double number, double multiple)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})"));
        }

        [Test]
        public void Multinomial()
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("MULTINOMIAL(2)"));
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3)"));
            Assert.AreEqual(1260, XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3,4)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("MULTINOMIAL(1E+100)"));
        }

        [Test]
        public void Multinomial_accepts_ranges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B2").InsertData(new[] { 2, 0, 5 });
            ws.Cell("A5").InsertData(new[] { 3, 6 });

            Assert.AreEqual(3087564480d, ws.Evaluate("MULTINOMIAL(B:XFD, 2, A5:A6)"));
        }

        [Test]
        public void Multinomial_doesnt_accept_negative_values()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("MULTINOMIAL(5, -1)"));
        }

        [Test]
        public void Multinomial_coercion()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = true;
            ws.Cell("A2").Value = 5;
            ws.Cell("A3").Value = "1 2/2";
            ws.Cell("A4").Value = "one";

            // True is not converted
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MULTINOMIAL(A1:A2)"));

            // Text is coerced
            Assert.AreEqual(21, ws.Evaluate("MULTINOMIAL(A2:A3)"));

            // Text is coerced, errors are propagates
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("MULTINOMIAL(A2:A4)"));

            // Errors are propagates
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("MULTINOMIAL(5, #DIV/0!)"));
        }

        [TestCase(1.5, ExpectedResult = 3)]
        [TestCase(3, ExpectedResult = 3)]
        [TestCase(2, ExpectedResult = 3)]
        [TestCase(-1, ExpectedResult = -1)]
        [TestCase(-2, ExpectedResult = -3)]
        [TestCase(0, ExpectedResult = 1)]
        [TestCase(1E+100, ExpectedResult = 1E+100)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double Odd(double number)
        {
            return (double)XLWorkbook.EvaluateExpr($"ODD({number})");
        }

        [Test]
        public void Pi()
        {
            Assert.AreEqual(Math.PI, XLWorkbook.EvaluateExpr("PI()"));
        }

        [TestCase(2, 3, ExpectedResult = 8)]
        [TestCase(2, 0.5, ExpectedResult = 1.414213562373)]
        [TestCase(-1.234, 5.0, ExpectedResult = -2.861381721051)]
        [TestCase(1.234, 5.1, ExpectedResult = 2.9221823578798)]
        [DefaultFloatingPointTolerance(1e-12)]
        public double Power(double x, double y)
        {
            return (double)XLWorkbook.EvaluateExpr($"POWER({x}, {y})");
        }

        [Test]
        public void Power_errors()
        {
            // Negative base and fractional exponent
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("POWER(-5, 0.5)"));

            // Spec says this should be #DIV/0!, but Excel says #NUM!
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("POWER(0, 0)"));

            // base is zero and exponent is negative -> #NUM!
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("POWER(0, -5)"));

            // Result is not representable (e.g. out fo range)
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("POWER(1e+100, 1e+100)"));
        }

        [Test]
        public void Product()
        {
            Assert.AreEqual(24d, XLWorkbook.EvaluateExpr("PRODUCT(2,3,4)"));

            // Examples from specification
            Assert.AreEqual(1d, XLWorkbook.EvaluateExpr("PRODUCT(1)"));
            Assert.AreEqual(120d, XLWorkbook.EvaluateExpr("PRODUCT(1,2,3,4,5)"));
            Assert.AreEqual(24d, XLWorkbook.EvaluateExpr("PRODUCT({1,2;3,4})"));
            Assert.AreEqual(120d, XLWorkbook.EvaluateExpr("PRODUCT({2,3},4,\"5\")"));

            // If no arguments are passed, return 0
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr("PRODUCT({\"hello\"})"));

            // Scalar blank is skipped
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("PRODUCT(IF(TRUE,), 1)"));

            // Scalar logical is converted to number
            Assert.AreEqual(0, XLWorkbook.EvaluateExpr("PRODUCT(FALSE, 1)"));
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("PRODUCT(2, TRUE)"));

            // Scalar text is converted to number
            Assert.AreEqual(5, XLWorkbook.EvaluateExpr("PRODUCT(\"5\")"));

            // Scalar text that is not convertible return error
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("PRODUCT(1, \"Hello\")"));

            // Array non-number arguments are ignored
            Assert.AreEqual(5, XLWorkbook.EvaluateExpr("PRODUCT({5, \"Hello\", FALSE, TRUE})"));

            // Reference argument only uses number, ignores blanks, logical and text
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = true;
            ws.Cell("A3").Value = "100";
            ws.Cell("A4").Value = "hello";
            ws.Cell("A5").Value = 2;
            ws.Cell("A6").Value = 3;
            Assert.AreEqual(6, ws.Evaluate("PRODUCT(A1:A6)"));

            // Scalar error is propagated
            Assert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("PRODUCT(1, #NULL!)"));

            // Array error is propagated
            Assert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("PRODUCT({1, #NULL!})"));

            // Reference error is propagated
            ws.Cell("A1").Value = XLError.NoValueAvailable;
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("PRODUCT(A1)"));
        }

        [TestCase(5, 2, ExpectedResult = 2)]
        [TestCase(4.5, 3.1, ExpectedResult = 1)]
        [TestCase(-10, 3, ExpectedResult = -3)]
        [TestCase(-10, -4, ExpectedResult = 2)]
        [TestCase(1E+100, 1E+40, ExpectedResult = 1E+60)]
        public double Quotient(double x, double y)
        {
            return (double)XLWorkbook.EvaluateExpr($"QUOTIENT({x}, {y})");
        }

        [Test]
        public void Quotient_errors()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("QUOTIENT(1, 0)"));
        }

        [TestCase(270, ExpectedResult = 4.71238898038469)]
        [TestCase(-180, ExpectedResult = -Math.PI)]
        [DefaultFloatingPointTolerance(XLHelper.Epsilon)]
        public double Radians(double angle)
        {
            return (double)XLWorkbook.EvaluateExpr($"RADIANS({angle})");
        }

        [Test]
        public void Rand()
        {
            for (var i = 0; i < 100; ++i)
            {
                var randomNumber = (double)XLWorkbook.EvaluateExpr("RAND()");
                Assert.That(randomNumber, Is.GreaterThanOrEqualTo(0).And.LessThan(1));
            }
        }

        [Test]
        public void RandBetween()
        {
            for (var i = 0; i < 100; ++i)
            {
                var randomNumber = (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(10, 20)");
                Assert.That(randomNumber, Is.GreaterThanOrEqualTo(10).And.LessThanOrEqualTo(20));
            }

            Assert.AreEqual(101, (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(100.5, 100.9)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("RANDBETWEEN(100.9, 100.5)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("RANDBETWEEN(20, 5)"));
            Assert.That((double)XLWorkbook.EvaluateExpr("RANDBETWEEN(1E+100, 1E+110)"), Is.GreaterThanOrEqualTo(1E+100).And.LessThanOrEqualTo(1E+110));
        }

        [TestCase(1, 0, ExpectedResult = "I")]
        [TestCase(3046, 1, ExpectedResult = @"MMMVLI")]
        [TestCase(3999, 1, ExpectedResult = @"MMMLMVLIV")]
        [TestCase(999, 0, ExpectedResult = @"CMXCIX")]
        [TestCase(999.99, 0.9, ExpectedResult = @"CMXCIX")]
        [TestCase(999, 1, ExpectedResult = @"LMVLIV")]
        [TestCase(999, 2, ExpectedResult = @"XMIX")]
        [TestCase(999, 3, ExpectedResult = @"VMIV")]
        [TestCase(999, 4, ExpectedResult = @"IM")]
        public string Roman(double value, double form)
        {
            return (string)XLWorkbook.EvaluateExpr($"ROMAN({value}, {form})");
        }

        [Test]
        public void Roman_value_0_is_empty_string()
        {
            Assert.AreEqual(string.Empty, XLWorkbook.EvaluateExpr("ROMAN(0, 0)"));
        }

        [Test]
        public void Roman_has_optional_second_argument_with_default_value_0()
        {
            Assert.AreEqual(@"CMXCIX", XLWorkbook.EvaluateExpr("ROMAN(999)"));
        }

        [Test]
        public void Roman_form_must_be_between_0_and_4()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(1, -1)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(1, 5)"));
        }

        [Test]
        public void Roman_value_must_be_between_0_and_3999()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(-1, 0)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROMAN(4000, 0)"));
        }

        [TestCase(2.15, 1, ExpectedResult = 2.2)]
        [TestCase(2.149, 1, ExpectedResult = 2.1)]
        [TestCase(-1.475, 2, ExpectedResult = -1.48)]
        [TestCase(21.5, -1, ExpectedResult = 20.0)]
        [TestCase(626.3, -3, ExpectedResult = 1000.0)]
        [TestCase(1.98, -1, ExpectedResult = 0.0)]
        [TestCase(-50.55, -2, ExpectedResult = -100.0)]
        [TestCase(31.565, 2, ExpectedResult = 31.57)]
        [TestCase(-31.565, 2, ExpectedResult = -31.57)]
        [TestCase(1E+100, 2, ExpectedResult = 1E+100)]
        [TestCase(1.25, 0, ExpectedResult = 1)]
        [TestCase(1, -1E+100, ExpectedResult = 0)]
        [TestCase(1.123456, 1E+100, ExpectedResult = 1.123456)] // Excel says 0 for anything over 2147483646
        public double Round(double number, double digits)
        {
            return (double)XLWorkbook.EvaluateExpr($"ROUND({number}, {digits})");
        }

        [TestCase(3.2, 0, ExpectedResult = 3.0)]
        [TestCase(76.9, 0, ExpectedResult = 76.0)]
        [TestCase(3.14159, 3, ExpectedResult = 3.141)]
        [TestCase(-3.14159, 1, ExpectedResult = -3.1)]
        [TestCase(31415.92654, -2, ExpectedResult = 31400.0)]
        [TestCase(0, 3, ExpectedResult = 0)]
        public double RoundDown(double number, double digits)
        {
            return (double)XLWorkbook.EvaluateExpr($"ROUNDDOWN({number}, {digits})");
        }

        [TestCase(3.2, 0, ExpectedResult = 4)]
        [TestCase(76.9, 0, ExpectedResult = 77.0)]
        [TestCase(3.14159, 3, ExpectedResult = 3.142)]
        [TestCase(-3.14159, 1, ExpectedResult = -3.2)]
        [TestCase(31415.92654, -2, ExpectedResult = 31500.0)]
        [TestCase(0, 3, ExpectedResult = 0)]
        [TestCase(11, 0, ExpectedResult = 11)]
        public double RoundUp(double number, double digits)
        {
            return (double)XLWorkbook.EvaluateExpr($"ROUNDUP({number}, {digits})");
        }

        [TestCase("0", ExpectedResult = 0)]
        [TestCase("10.5", ExpectedResult = 1)]
        [TestCase("-5.4", ExpectedResult = -1)]
        [TestCase("-0.00001", ExpectedResult = -1)]
        [TestCase("-1E+300", ExpectedResult = -1)]
        [TestCase("\"0 1/2\"", ExpectedResult = 1)]
        [TestCase("FALSE", ExpectedResult = 0)]
        [TestCase("TRUE", ExpectedResult = 1)]
        public double Sign(string arg)
        {
            return (double)XLWorkbook.EvaluateExpr($"SIGN({arg})");
        }

        [TestCase("0", ExpectedResult = 0)]
        [TestCase("1", ExpectedResult = 0.8414709848078965)]
        [TestCase("-1", ExpectedResult = -0.8414709848078965)]
        [TestCase("PI()", ExpectedResult = 0)]
        [TestCase("PI()/2", ExpectedResult = 1)]
        [TestCase("30*PI()/180", ExpectedResult = 0.5)]
        [TestCase("RADIANS(30)", ExpectedResult = 0.5)]
        [DefaultFloatingPointTolerance(tolerance)]
        public double Sin(string arg)
        {
            return (double)XLWorkbook.EvaluateExpr($"SIN({arg})");
        }

        [TestCase("0", 0)]
        [TestCase("1", 1.1752011936438014)]
        [TestCase("10", 11013.232874703393)]
        [TestCase("100", 1.3440585709080678E+43)]
        [TestCase("100", 1.3440585709080678E+43)]
        [TestCase("711", XLError.NumberInvalid)]
        [TestCase("-711", XLError.NumberInvalid)]
        public void Sinh(string arg, object result)
        {
            var actual = XLWorkbook.EvaluateExpr($"SINH({arg})");
            Assert.AreEqual(result, actual);
        }

        [TestCase(0, 1)]
        [TestCase(0.3, 1.0467516)]
        [TestCase(0.6, 1.21162831)]
        [TestCase(0.9, 1.60872581)]
        [TestCase(1.2, 2.759703601)]
        [TestCase(1.5, 14.1368329)]
        [TestCase(1.8, -4.401367872)]
        [TestCase(2.1, -1.980801656)]
        [TestCase(2.4, -1.356127641)]
        [TestCase(2.7, -1.10610642)]
        [TestCase(3.0, -1.010108666)]
        [TestCase(3.3, -1.012678974)]
        [TestCase(3.6, -1.115127532)]
        [TestCase(3.9, -1.377538917)]
        [TestCase(4.2, -2.039730601)]
        [TestCase(4.5, -4.743927548)]
        [TestCase(4.8, 11.42870421)]
        [TestCase(5.1, 2.645658426)]
        [TestCase(5.4, 1.575565187)]
        [TestCase(5.7, 1.198016873)]
        [TestCase(6.0, 1.041481927)]
        [TestCase(6.3, 1.000141384)]
        [TestCase(6.6, 1.052373922)]
        [TestCase(6.9, 1.225903187)]
        [TestCase(7.2, 1.643787029)]
        [TestCase(7.5, 2.884876262)]
        [TestCase(7.8, 18.53381902)]
        [TestCase(8.1, -4.106031636)]
        [TestCase(8.4, -1.925711244)]
        [TestCase(8.7, -1.335743646)]
        [TestCase(9.0, -1.097537906)]
        [TestCase(9.3, -1.007835594)]
        [TestCase(9.6, -1.015550252)]
        [TestCase(9.9, -1.124617578)]
        [TestCase(10.2, -1.400039323)]
        [TestCase(10.5, -2.102886109)]
        [TestCase(10.8, -5.145888341)]
        [TestCase(11.1, 9.593612018)]
        [TestCase(11.4, 2.541355049)]
        [TestCase(45, 1.90359)]
        [TestCase(30, 6.48292)]
        public void Sec_returns_correct_number(double angle, double expectedOutput)
        {
            var result = (double)XLWorkbook.EvaluateExpr($"SEC({angle})");
            Assert.AreEqual(expectedOutput, result, 0.00001);

            // as the secant is symmetric for positive and negative numbers, let's assert twice:
            var resultForNegative = (double)XLWorkbook.EvaluateExpr($"SEC({-angle})");
            Assert.AreEqual(expectedOutput, resultForNegative, 0.00001);
        }

        [TestCase(-9, 0.00024682)]
        [TestCase(-8, 0.000670925)]
        [TestCase(-7, 0.001823762)]
        [TestCase(-6, 0.004957474)]
        [TestCase(-5, 0.013475282)]
        [TestCase(-4, 0.036618993)]
        [TestCase(-3, 0.099327927)]
        [TestCase(-2, 0.265802229)]
        [TestCase(-1, 0.648054274)]
        [TestCase(0, 1)]
        [TestCase(1E+100, 0)]
        [TestCase(1E-100, 1)]
        public void Sech_returns_correct_number(double angle, double expectedOutput)
        {
            var result = (double)XLWorkbook.EvaluateExpr($"SECH({angle})");
            Assert.AreEqual(expectedOutput, result, 0.00001);

            // as the secant is symmetric for positive and negative numbers, let's assert twice:
            var resultForNegative = (double)XLWorkbook.EvaluateExpr($"SECH({-angle})");
            Assert.AreEqual(expectedOutput, resultForNegative, 0.00001);
        }

        [Test]
        [DefaultFloatingPointTolerance(tolerance)]
        public void SeriesSum()
        {
            Assert.AreEqual(40.0, XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)"));

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A2").FormulaA1 = "PI()/4";
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
            ws.Cell("A5").FormulaA1 = "1/FACT(4)";
            ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

            var actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
            Assert.AreEqual(0.70710321482284566, actual);
        }

        [TestCase("{1,2,3;4,5,6}")]
        [TestCase("{1,2,3,4,5,6}")]
        [TestCase("{1,2;3,4;5,6}")]
        public void SeriesSum_takes_coefficients_row_by_row_left_to_right(string array)
        {
            Assert.AreEqual(1284, XLWorkbook.EvaluateExpr($"SERIESSUM(2,2,1,{array})"));
        }

        [Test]
        public void SeriesSum_returns_invalid_number_error_when_result_is_too_large()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[] { 1, 2, 3, 4, 5 });
            Assert.AreEqual(3E+300, ws.Evaluate("SERIESSUM(10,100,100,A1:A3)"));
            Assert.AreEqual(XLError.NumberInvalid, ws.Evaluate("SERIESSUM(10,100,100,A1:A4)"));
        }

        [Test]
        public void SeriesSum_coercion()
        {
            // For some weird reason, SERIESSUM doesn't convert logical
            foreach (var invalidValue in new[] { "\"\"", "TRUE" })
            {
                Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"SERIESSUM({invalidValue},1,1,1)"));
                Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"SERIESSUM(1,{invalidValue},1,1)"));
                Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,{invalidValue},1)"));
                Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,1,{invalidValue})"));
            }

            // Blank and text values are coerced to a number
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            foreach (var validArg in new[] { "A1", "\"0 0/2\"" })
            {
                Assert.AreEqual(0, ws.Evaluate($"SERIESSUM({validArg},1,1,1)"));
                Assert.AreEqual(1, ws.Evaluate($"SERIESSUM(1,{validArg},1,1)"));
                Assert.AreEqual(1, ws.Evaluate($"SERIESSUM(1,1,{validArg},1)"));
            }

            // Text is not converted in an area and causes conversion error
            ws.Cell("B2").Value = "0";
            ws.Cell("B3").Value = 5;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SERIESSUM(1,1,1,B2:B3)"));

            // Blank is interpreted as 0
            ws.Cell("C1").Value = Blank.Value;
            ws.Cell("C2").Value = 2;
            Assert.AreEqual(2, ws.Evaluate("SERIESSUM(1,1,1,C1:C2)"));
        }

        [TestCase(0, 0)]
        [TestCase(1, 1)]
        [TestCase(2, 1.4142135624)]
        [TestCase(1E+300, 1E+150)]
        public void Sqrt(double x, double result)
        {
            Assert.AreEqual(result, (double)XLWorkbook.EvaluateExpr($"SQRT({x})"), tolerance);
        }

        [TestCase(-1)]
        [TestCase(-0.0001)]
        public void Sqrt_returns_invalid_number_for_negative_numbers(double x)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"SQRT({x})"));
        }

        [Test]
        public void SqrtPi()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(1)");
            Assert.AreEqual(1.7724538509055159, actual, tolerance);

            actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(2)");
            Assert.AreEqual(2.5066282746310002, actual, tolerance);

            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("SQRTPI(-1)"));
        }

        [Test]
        public void Subtotal()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Non-existent functions return error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(0, A1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(0.9, A1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(12, A1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(100.9, A1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUBTOTAL(112, A1)"));
        }

        [Test]
        public void SubtotalAverage()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(1,A1,A2)";
            ws.Cell("A4").Value = "A";

            Assert.AreEqual(2.5, ws.Cell("A3").Value);
            Assert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(1, A1:A4)"));

            ws.Row(2).Hide();
            Assert.AreEqual(2, ws.Evaluate("SUBTOTAL(101, A1:A4)"));
        }

        [Test]
        public void Subtotal10Calc()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.DefinedNames.Add("subtotalrange", "$A$37:$A$38");

            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 4;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)"; // simple add subtotal
            ws.Cell("A4").Value = 8;
            ws.Cell("A5").Value = 16;
            ws.Cell("A6").FormulaA1 = "SUBTOTAL(9, A4:A5)"; // simple add subtotal
            ws.Cell("A7").Value = 32;
            ws.Cell("A8").Value = 64;
            ws.Cell("A9").FormulaA1 = "SUM(A7:A8)"; // func but not subtotal
            ws.Cell("A10").Value = 128;
            ws.Cell("A11").Value = 256;
            ws.Cell("A12").FormulaA1 = "SUBTOTAL(1, A10:A11)"; // simple avg subtotal
            ws.Cell("A13").Value = 512;
            ws.Cell("A14").FormulaA1 = "SUBTOTAL(9, A1:A13)"; // subtotals in range
            ws.Cell("A15").Value = 1024;
            ws.Cell("A16").Value = 2048;
            ws.Cell("A17").FormulaA1 = "42 + SUBTOTAL(9, A15:A16)"; // simple add subtotal in formula
            ws.Cell("A18").Value = 4096;
            ws.Cell("A19").FormulaA1 = "SUBTOTAL(9, A15:A18)"; // subtotals in range
            ws.Cell("A20").Value = 8192;
            ws.Cell("A21").Value = 16384;
            ws.Cell("A22").FormulaA1 = @"32768 * SEARCH(""SUBTOTAL(9, A1:A2)"", A28)"; // subtotal literal in formula
            ws.Cell("A23").FormulaA1 = "SUBTOTAL(9, A20:A22)"; // subtotal literal in formula in range
            ws.Cell("A24").Value = 65536;
            ws.Cell("A25").FormulaA1 = "A23"; // link to subtotal
            ws.Cell("A26").FormulaA1 = "PRODUCT(SUBTOTAL(9, A24:A25), 2)"; // subtotal as parameter in func
            ws.Cell("A27").Value = 131072;
            ws.Cell("A28").Value = "SUBTOTAL(9, A1:A2)"; // subtotal literal
            ws.Cell("A29").FormulaA1 = "SUBTOTAL(9, A27:A28)"; // subtotal literal in range
            ws.Cell("A30").FormulaA1 = "SUBTOTAL(9, A31:A32)"; // simple add subtotal backward
            ws.Cell("A31").Value = 262144;
            ws.Cell("A32").Value = 524288;
            ws.Cell("A33").FormulaA1 = "SUBTOTAL(9, A20:A32)"; // subtotals in range
            ws.Cell("A34").FormulaA1 = @"SUBTOTAL(VALUE(""9""), A1:A33, A35:A41)"; // func as parameter in subtotal and many ranges
            ws.Cell("A35").Value = 1048576;
            ws.Cell("A36").FormulaA1 = "SUBTOTAL(9, A31:A32, A35)"; // many ranges
            ws.Cell("A37").Value = 2097152;
            ws.Cell("A38").Value = 4194304;
            ws.Cell("A39").FormulaA1 = "SUBTOTAL(3*3, subtotalrange)"; // formula as parameter in subtotal and named range
            ws.Cell("A40").Value = 8388608;
            ws.Cell("A41").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2)"; // formula with link as parameter in subtotal
            ws.Cell("A42").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2) + SUBTOTAL(A4+1, A35:A40)"; // two subtotals in one formula

            Assert.AreEqual(6, ws.Cell("A3").Value);
            Assert.AreEqual(24, ws.Cell("A6").Value);
            Assert.AreEqual(192, ws.Cell("A12").Value);
            Assert.AreEqual(1118, ws.Cell("A14").Value);
            Assert.AreEqual(3114, ws.Cell("A17").Value);
            Assert.AreEqual(7168, ws.Cell("A19").Value);
            Assert.AreEqual(57344, ws.Cell("A23").Value);
            Assert.AreEqual(245760, ws.Cell("A26").Value);
            Assert.AreEqual(131072, ws.Cell("A29").Value);
            Assert.AreEqual(786432, ws.Cell("A30").Value);
            Assert.AreEqual(1097728, ws.Cell("A33").Value);
            Assert.AreEqual(16834654, ws.Cell("A34").Value);
            Assert.AreEqual(1835008, ws.Cell("A36").Value);
            Assert.AreEqual(6291456, ws.Cell("A39").Value);
            Assert.AreEqual(31457280, ws.Cell("A41").Value);
            Assert.AreEqual(47185920, ws.Cell("A42").Value);
        }

        [Test]
        public void Subtotal100Calc()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = 2;
            ws.Cell("C1").Value = Blank.Value;
            ws.Cell("A2").Value = "A";
            ws.Cell("B2").Value = 4;
            ws.Cell("C2").Value = 8;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(109, A1:A2)";
            ws.Cell("B3").FormulaA1 = "SUBTOTAL(109, B1:B2)";
            ws.Cell("C3").FormulaA1 = "SUBTOTAL(109, C1:C2)";
            ws.Cell("A4").Value = 16;
            ws.Cell("B4").Value = 32;
            ws.Cell("C4").Value = 64;
            ws.Cell("A5").Value = 128;
            ws.Cell("B5").Value = 256;
            ws.Cell("C5").Value = 512;
            ws.Cell("A6").FormulaA1 = "SUBTOTAL(109, A1:A5)";
            ws.Cell("B6").FormulaA1 = "SUBTOTAL(109, B1:B5)";
            ws.Cell("C6").FormulaA1 = "SUBTOTAL(109, C1:C5)";

            ws.Row(2).Hide();
            ws.Row(5).Hide();

            Assert.AreEqual(1, ws.Cell("A3").Value);
            Assert.AreEqual(2, ws.Cell("B3").Value);
            Assert.AreEqual(0, ws.Cell("C3").Value);
            Assert.AreEqual(17, ws.Cell("A6").Value);
            Assert.AreEqual(34, ws.Cell("B6").Value);
            Assert.AreEqual(64, ws.Cell("C6").Value);
        }

        [Test]
        public void SubtotalCount()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(2,A1:A3)";

            Assert.AreEqual(2, ws.Cell("A4").Value);
            Assert.AreEqual(1, ws.Evaluate("SUBTOTAL(2,A2:A4)"));

            ws.Row(2).Hide();
            Assert.AreEqual(1, ws.Evaluate("SUBTOTAL(102,A1:A4)"));
        }

        [Test]
        public void SubtotalCountA()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = string.Empty;
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(3,A1,A2,A3)";

            Assert.AreEqual(3, ws.Cell("A4").Value);
            Assert.AreEqual(3, ws.Evaluate("SUBTOTAL(3,A1:A4)"));

            ws.Row(1).Hide();
            Assert.AreEqual(2, ws.Evaluate("SUBTOTAL(103,A1:A4)"));
        }

        [Test]
        public void SubtotalMax()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(4,A1,A2,A3) + 10";

            Assert.AreEqual(13, ws.Cell("A4").Value);
            Assert.AreEqual(3, ws.Evaluate("SUBTOTAL(4,A1:A4)"));

            ws.Cell("A5").Value = 2.5;
            ws.Row(2).Hide();
            Assert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(104,A1:A5)"));
        }

        [Test]
        public void SubtotalMin()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(5,A1,A2,A3) - 10";

            Assert.AreEqual(-8, ws.Cell("A4").Value);
            Assert.AreEqual(2, ws.Evaluate("SUBTOTAL(5,A1:A4)"));

            ws.Cell("A5").Value = 2.5;
            ws.Row(1).Hide();
            Assert.AreEqual(2.5, ws.Evaluate("SUBTOTAL(105,A1:A5)"));
        }

        [Test]
        public void SubtotalProduct()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(6,A1,A2,A3)";

            Assert.AreEqual(6, ws.Cell("A4").Value);
            Assert.AreEqual(6, ws.Evaluate("SUBTOTAL(6,A1:A4)"));

            ws.Row(2).Hide();
            ws.Cell("A5").Value = 4;
            Assert.AreEqual(8, ws.Evaluate("SUBTOTAL(106,A1:A5)"));
        }

        [Test]
        [DefaultFloatingPointTolerance(XLHelper.Epsilon)]
        public void SubtotalStDev()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(7,A1:A3,A5)";
            ws.Cell("A5").Value = 5;

            Assert.AreEqual(1.5275252316, (double)ws.Cell("A4").Value);
            Assert.AreEqual(1.5275252316, (double)ws.Evaluate("SUBTOTAL(7,A1:A5)"));

            ws.Row(2).Hide();
            Assert.AreEqual(2.1213203435, (double)ws.Evaluate("SUBTOTAL(107,A1:A5)"));
        }

        [Test]
        public void SubtotalStDevP()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(8,A1,A2,A3)";

            Assert.AreEqual(0.5, ws.Cell("A4").Value);
            Assert.AreEqual(0.5, ws.Evaluate("SUBTOTAL(8,A1:A4)"));

            ws.Row(2).Hide();
            ws.Cell("A5").Value = 3;
            Assert.AreEqual(0.5, ws.Evaluate("SUBTOTAL(108,A1:A5)"));
        }

        [Test]
        public void SubtotalSum()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(9,A1,A2,A3)";

            Assert.AreEqual(5, ws.Cell("A4").Value);
            Assert.AreEqual(5, ws.Evaluate("SUBTOTAL(9,A1:A4)"));

            ws.Row(2).Hide();

            Assert.AreEqual(2, ws.Evaluate("SUBTOTAL(109, A1:A4)"));
        }

        [Test]
        public void SubtotalVar()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 5;
            ws.Cell("A2").Value = 4;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").Value = 8;
            ws.Cell("A5").Value = 5;
            ws.Cell("A6").FormulaA1 = "SUBTOTAL(10,A1:A5)";

            Assert.AreEqual(3, ws.Cell("A6").Value);
            Assert.AreEqual(3, ws.Evaluate("SUBTOTAL(10,A1:A6)"));

            ws.Row(1).Hide();
            ws.Row(5).Hide();
            Assert.AreEqual(8, ws.Evaluate("SUBTOTAL(110,A1:A6)"));
        }

        [Test]
        public void SubtotalVarP()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 3;
            ws.Cell("A3").Value = "A";
            ws.Cell("A4").FormulaA1 = "SUBTOTAL(11,A1,A2,A3)";

            Assert.AreEqual(0.25, ws.Cell("A4").Value);
            Assert.AreEqual(0.25, ws.Evaluate("SUBTOTAL(11,A1:A4)"));

            ws.Row(2).Hide();
            ws.Cell("A5").Value = 4;
            Assert.AreEqual(1, ws.Evaluate("SUBTOTAL(111,A1:A5)"));
        }

        [Test]
        public void Sum()
        {
            IXLCell cell = new XLWorkbook().AddWorksheet("Sheet1").FirstCell();
            IXLCell fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
            fCell.FormulaA1 = "sum(A1:A2)";

            Assert.AreEqual(3.0, fCell.Value);
        }

        [Test]
        public void SumDateTimeAndNumber()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = 1;
                ws.Cell("A2").Value = new DateTime(2018, 1, 1);
                Assert.AreEqual(43102, ws.Evaluate("SUM(A1:A2)"));

                ws.Cell("A1").Value = 2;
                ws.Cell("A2").FormulaA1 = "DATE(2018,1,1)";
                Assert.AreEqual(43103, ws.Evaluate("SUM(A1:A2)"));
            }
        }

        [TestCase(9, "SUMIF(A:B, \"A*\", C:C)")]
        [TestCase(9, "SUMIF(A1:B6, \"A*\", C1:C6)")]
        public void SumIf_InputRangeHasMultipleColumns(int expectedOutcome, string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            var data = new object[]
            {
                    new { Id = "AA", Id2 = "BA", Value = 2},
                    new { Id = "AB", Id2 = "BB", Value = 3},
                    new { Id = "BA", Id2 = "AA", Value = 2},
                    new { Id = "BB", Id2 = "AB", Value = 1},
                    new { Id = "AC", Id2 = "AC", Value = 4},
            };
            ws.Cell("A1").InsertTable(data);

            Assert.AreEqual(expectedOutcome, ws.Evaluate(formula));
        }

        /// <summary>
        /// refers to Example 1 from the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(63000, "SUMIF(A1:A4,\">160000\", B1:B4)")]
        [TestCase(900000, "SUMIF(A1:A4,\">160000\")")]
        [TestCase(21000, "SUMIF(A1:A4, 300000, B1:B4)")]
        [TestCase(28000, "SUMIF(A1:A4, \">\" &C1, B1:B4)")]
        public void SumIf_ReturnsCorrectValues_ReferenceExample1FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(1, 1).Value = 100000;
                ws.Cell(1, 2).Value = 7000;
                ws.Cell(2, 1).Value = 200000;
                ws.Cell(2, 2).Value = 14000;
                ws.Cell(3, 1).Value = 300000;
                ws.Cell(3, 2).Value = 21000;
                ws.Cell(4, 1).Value = 400000;
                ws.Cell(4, 2).Value = 28000;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to Example 2 from the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(2000, "SUMIF(A2:A7,\"Fruits\", C2:C7)")]
        [TestCase(12000, "SUMIF(A2:A7,\"Vegetables\", C2:C7)")]
        [TestCase(4300, "SUMIF(B2:B7, \"*es\", C2:C7)")]
        [TestCase(400, "SUMIF(A2:A7, \"\", C2:C7)")]
        public void SumIf_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                wb.ReferenceStyle = XLReferenceStyle.A1;

                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(2, 1).Value = "Vegetables";
                ws.Cell(3, 1).Value = "Vegetables";
                ws.Cell(4, 1).Value = "Fruits";
                ws.Cell(5, 1).Value = "";
                ws.Cell(6, 1).Value = "Vegetables";
                ws.Cell(7, 1).Value = "Fruits";

                ws.Cell(2, 2).Value = "Tomatoes";
                ws.Cell(3, 2).Value = "Celery";
                ws.Cell(4, 2).Value = "Oranges";
                ws.Cell(5, 2).Value = "Butter";
                ws.Cell(6, 2).Value = "Carrots";
                ws.Cell(7, 2).Value = "Apples";

                ws.Cell(2, 3).Value = 2300;
                ws.Cell(3, 3).Value = 5500;
                ws.Cell(4, 3).Value = 800;
                ws.Cell(5, 3).Value = 400;
                ws.Cell(6, 3).Value = 4200;
                ws.Cell(7, 3).Value = 1200;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        [Test]
        public void SumIf_ReturnsCorrectValues_WhenCalledOnFullColumn()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                var data = new object[]
                {
                    new { Id = "A", Value = 2},
                    new { Id = "B", Value = 3},
                    new { Id = "C", Value = 2},
                    new { Id = "A", Value = 1},
                    new { Id = "B", Value = 4}
                };
                ws.Cell("A1").InsertTable(data);
                var formula = "=SUMIF(A:A,\"=A\",B:B)";
                var value = ws.Evaluate(formula);
                Assert.AreEqual(3, value);
            }
        }

        [Test]
        public void SumIf_ReturnsCorrectValues_WhenFormulaBelongToSameRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                var data = new object[]
                {
                    new { Id = "A", Value = 2},
                    new { Id = "B", Value = 3},
                    new { Id = "C", Value = 2},
                    new { Id = "A", Value = 1},
                    new { Id = "B", Value = 4},
                };
                ws.Cell("A1").InsertTable(data);
                ws.Cell("A7").SetValue("Sum A");
                // SUMIF formula
                var formula = "=SUMIF(A:A,\"=A\",B:B)";
                ws.Cell("B7").SetFormulaA1(formula);
                var value = ws.Cell("B7").Value;
                Assert.AreEqual(3, value);
            }
        }

        [Test]
        public void SumIfs_MultidimensionalRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertData(new object[]
            {
                (10, 10, 1, 2),
                (20, 15, 2, 4),
                (30, 20, 3, 6),
                (40, 25, 4, 8),
                (50, 30, 5, 10),
            });
            Assert.AreEqual(30, ws.Evaluate("SUMIFS(C1:D5,A1:B5,\">20\")"));
        }

        /// <summary>
        /// refers to Example 2 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedResult"></param>
        /// <param name="formula"></param>
        [TestCase(2000, "SUMIFS(C2:C7, A2:A7, \"Fruits\")")]
        [TestCase(12000, "SUMIFS(C2:C7, A2:A7, \"Vegetables\")")]
        [TestCase(4300, "SUMIFS(C2:C7, B2:B7, \"*es\")")]
        [TestCase(400, "SUMIFS(C2:C7, A2:A7, \"\")")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedResult, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(2, 1).Value = "Vegetables";
                ws.Cell(3, 1).Value = "Vegetables";
                ws.Cell(4, 1).Value = "Fruits";
                ws.Cell(5, 1).Value = "";
                ws.Cell(6, 1).Value = "Vegetables";
                ws.Cell(7, 1).Value = "Fruits";

                ws.Cell(2, 2).Value = "Tomatoes";
                ws.Cell(3, 2).Value = "Celery";
                ws.Cell(4, 2).Value = "Oranges";
                ws.Cell(5, 2).Value = "Butter";
                ws.Cell(6, 2).Value = "Carrots";
                ws.Cell(7, 2).Value = "Apples";

                ws.Cell(2, 3).Value = 2300;
                ws.Cell(3, 3).Value = 5500;
                ws.Cell(4, 3).Value = 800;
                ws.Cell(5, 3).Value = 400;
                ws.Cell(6, 3).Value = 4200;
                ws.Cell(7, 3).Value = 1200;

                ws.Cell(1, 3).Value = 300000;

                var actualResult = (double)ws.Evaluate(formula);
                Assert.AreEqual(expectedResult, actualResult);
            }
        }

        /// <summary>
        /// refers to Example 1 to SumIf from the Excel documentation.
        /// As SumIfs should behave the same if called with three parameters, but in a different order
        /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        /// <param name="expectedOutcome"></param>
        /// <param name="formula"></param>
        [TestCase(63000, "SUMIFS(B1:B4, A1:A4, \">160000\")")]
        [TestCase(21000, "SUMIFS(B1:B4, A1:A4, 300000)")]
        [TestCase(28000, "SUMIFS(B1:B4, A1:A4, \">\" &C1)")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExampleForSumIf1FromMicrosoft(int expectedOutcome, string formula)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(1, 1).Value = 100000;
                ws.Cell(1, 2).Value = 7000;
                ws.Cell(2, 1).Value = 200000;
                ws.Cell(2, 2).Value = 14000;
                ws.Cell(3, 1).Value = 300000;
                ws.Cell(3, 2).Value = 21000;
                ws.Cell(4, 1).Value = 400000;
                ws.Cell(4, 2).Value = 28000;

                ws.Cell(1, 3).Value = 300000;

                Assert.AreEqual(expectedOutcome, (double)ws.Evaluate(formula));
            }
        }

        /// <summary>
        /// refers to example data and formula to SumIfs in the Excel documentation,
        /// <see cref="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
        /// </summary>
        [TestCase(20, "=SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")")]
        [TestCase(30, "=SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")")]
        public void SumIfs_ReturnsCorrectValues_ReferenceExampleFromMicrosoft(
            int expectedResult,
            string formula)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var row = 2;

                ws.Cell(row, 1).Value = 5;
                ws.Cell(row, 2).Value = "Apples";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 4;
                ws.Cell(row, 2).Value = "Apples";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 15;
                ws.Cell(row, 2).Value = "Artichokes";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 3;
                ws.Cell(row, 2).Value = "Artichokes";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 22;
                ws.Cell(row, 2).Value = "Bananas";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 12;
                ws.Cell(row, 2).Value = "Bananas";
                ws.Cell(row, 3).Value = "Sarah";
                row++;

                ws.Cell(row, 1).Value = 10;
                ws.Cell(row, 2).Value = "Carrots";
                ws.Cell(row, 3).Value = "Tom";
                row++;

                ws.Cell(row, 1).Value = 33;
                ws.Cell(row, 2).Value = "Carrots";
                ws.Cell(row, 3).Value = "Sarah";

                var actualResult = ws.Evaluate(formula);

                Assert.AreEqual(expectedResult, (double)actualResult, tolerance);
            }
        }

        [TestCase("SUMIFS(D1:E5,A1:B5,\"A*\",C1:C5,\">2\")")]
        [TestCase("SUMIFS(H1:I3,A1:B3,1,D1:F2,2)")]
        [TestCase("SUMIFS(D:E,A:B,\"A*\",C:C,\">2\")")]
        public void SumIfs_ReturnsErrorWhenRangeDimensionsAreNotSame(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula));
        }

        [TestCase("SUMIFS(A1:A3, B1:B3,\"<>B\")", 11)]
        [TestCase("SUMIFS(A1:A3, B1:B3,\"<>\")", 110)]
        public void SumIfs_matches_blank_cells_when_criteria_is_not_equal(string formula, double expectedSum)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 10;
            ws.Cell("A3").Value = 100;
            ws.Cell("B1").Value = Blank.Value;
            ws.Cell("B2").Value = string.Empty;
            ws.Cell("B3").Value = "B";

            Assert.AreEqual(expectedSum, ws.Evaluate(formula));
        }

        [Test]
        public void SumProduct()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().InsertData(Enumerable.Range(1, 10));
            ws.FirstCell().CellRight().InsertData(Enumerable.Range(1, 10).Reverse());

            Assert.AreEqual(2, ws.Evaluate("SUMPRODUCT(A2)"));
            Assert.AreEqual(55, ws.Evaluate("SUMPRODUCT(A1:A10)"));
            Assert.AreEqual(220, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B5)"));

            // Scalar, one element array and single cell area are compatible
            Assert.AreEqual(60, ws.Evaluate("SUMPRODUCT(A5, 4, {3})"));

            // An array can be an argument
            Assert.AreEqual(10, ws.Evaluate("SUMPRODUCT(A1:A3, {3;2;1})"));

            // An array must have correct orientation, otherwise dimensions don't match
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(A1:A3, {3,2,1})"));

            // Anything but number is counted as zero. The second array is zero for all values = result is 0.
            Assert.AreEqual(0, ws.Evaluate("SUMPRODUCT({1,2,3,4}, {TRUE,FALSE,\"1\",\"\"})"));

            // Any error returns error
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMPRODUCT({1,2}, {1,#N/A})"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMPRODUCT(A1, #N/A)"));
            ws.Cell("A2").Value = XLError.NoValueAvailable;
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMPRODUCT(A2, 5)"));

            // Blank cells and cells with text should be treated as zeros
            ws.Range("A1:A5").Clear();
            Assert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

            // Non-number values are treated as zero
            ws.Range("A1:A5").SetValue("asdf");
            Assert.AreEqual(110, ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)"));

            // Blank cell is considered as a blank and cause #VALUE! error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(Z1, 5)"));

            // Blank value will cause #VALUE! error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("SUMPRODUCT(IF(TRUE,,), 5)"));
        }

        [Test]
        public void SumSq()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Examples from specification
            Assert.AreEqual(4.0, XLWorkbook.EvaluateExpr("SUMSQ(2)"));
            Assert.AreEqual(19.21, XLWorkbook.EvaluateExpr("SUMSQ(2.5, -3.6)"));
            Assert.AreEqual(24.97, XLWorkbook.EvaluateExpr("SUMSQ({ 2.5, -3.6}, 2.4)"));

            // Scalar blank is converted to 0
            Assert.AreEqual(16, XLWorkbook.EvaluateExpr("SUMSQ(IF(TRUE,), 4)"));

            // Scalar logical is converted to number
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("SUMSQ(3, TRUE)"));

            // Scalar text is converted to number
            Assert.AreEqual(25, XLWorkbook.EvaluateExpr("SUMSQ(\"4\", \"3\")"));

            // Scalar text that is not convertible return error
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("SUMSQ(1, \"Hello\")"));

            // Array logical arguments are ignored
            Assert.AreEqual(4, XLWorkbook.EvaluateExpr("SUMSQ({2,TRUE,TRUE,FALSE,FALSE})"));

            // Array text arguments are ignored
            Assert.AreEqual(20, XLWorkbook.EvaluateExpr("SUMSQ({4, 2, \"hello\", \"10\" })"));

            // Blank, logical and text from reference are ignored
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").Value = true;
            ws.Cell("A3").Value = "100";
            ws.Cell("A4").Value = "hello";
            ws.Cell("A5").Value = 1;
            ws.Cell("A6").Value = 4;
            Assert.AreEqual(17, ws.Evaluate("SUMSQ(A1:A6)"));

            // Scalar error is propagated
            Assert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("SUMSQ(1, #NULL!)"));

            // Array error is propagated
            Assert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("SUMSQ({1, #NULL!})"));

            // Reference error is propagated
            ws.Cell("A1").Value = XLError.NoValueAvailable;
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("SUMSQ(A1)"));
        }

        [TestCase(-1, ExpectedResult = -1.5574077247)]
        [TestCase(0, ExpectedResult = 0)]
        [TestCase(1, ExpectedResult = 1.5574077247)]
        [TestCase(134217727, ExpectedResult = 3.2584564256)]
        [TestCase(-134217727, ExpectedResult = -3.2584564256)]
        [DefaultFloatingPointTolerance(tolerance)]
        public double Tan(double radians)
        {
            return (double)XLWorkbook.EvaluateExpr($"TAN({radians})");
        }

        [TestCase(134217728)]
        [TestCase(-134217728)]
        [TestCase(1E+100)]
        public void Tan_returns_invalid_number_for_radians_outside_limit(double radians)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"TAN({radians})"));
        }

        [TestCase(-1, -0.761594156)]
        [TestCase(0, 0)]
        [TestCase(1, 0.761594156)]
        [TestCase(1E+300, 1)]
        [TestCase(-1E+300, -1)]
        [DefaultFloatingPointTolerance(tolerance)]
        public void Tanh(double number, double result)
        {
            Assert.AreEqual(result, (double)XLWorkbook.EvaluateExpr($"TANH({number})"));
        }

        [TestCase(27.64799257, null, 27)]
        [TestCase(0, null, 0)]
        [TestCase(0, 0, 0)]
        [TestCase(3.1415926, 0, 3)]
        [TestCase(3.1415926, 1, 3.1)]
        [TestCase(3.1415926, 3, 3.141)]
        [TestCase(3.1415926, 5, 3.14159)]
        [TestCase(-4.3, 0, -4)]
        [TestCase(8.9, null, 8)]
        [TestCase(-8.9, null, -8)]
        [TestCase(0.45, null, 0)]
        public void Trunc(double number, double? digits, object expectedResult)
        {
            var formula = digits is null ? $"TRUNC({number})" : $"TRUNC({number}, {digits})";
            Assert.AreEqual(expectedResult, (double)XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase(27.64799257, -1, 20)]
        [TestCase(27.64799257, 0, 27)]
        [TestCase(27.64799257, 1, 27.6)]
        [TestCase(27.64799257, 4, 27.6479)]
        public void Trunc_Specify_Digits(double input, int digits, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"TRUNC({input.ToString(CultureInfo.InvariantCulture)}, {digits})");
            Assert.AreEqual(expectedResult, actual);
        }
    }
}

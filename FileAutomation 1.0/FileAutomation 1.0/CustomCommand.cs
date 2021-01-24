using System.Windows.Input;

namespace FileAutomation_1._0
{
    public class CustomCommand
    {
        public static readonly RoutedUICommand FirstPage = new RoutedUICommand
                (
                        "FirstPage",
                        "FirstPage",
                        typeof(CustomCommand),
                        null
                );

        public static readonly RoutedUICommand LastPage = new RoutedUICommand
                (
                        "LastPage",
                        "LastPage",
                        typeof(CustomCommand),
                        null
                );


        public static readonly RoutedUICommand PreviousPage = new RoutedUICommand
                (
                        "PreviousPage",
                        "PreviousPage",
                        typeof(CustomCommand),
                        null
                );


        public static readonly RoutedUICommand NextPage = new RoutedUICommand
                (
                        "NextPage",
                        "NextPage",
                        typeof(CustomCommand),
                        null
                );

        public static readonly RoutedUICommand Search = new RoutedUICommand
            (
                    "Search",
                    "Search",
                    typeof(CustomCommand),
                    null
            );
        public static readonly RoutedUICommand SearchNext = new RoutedUICommand
            (
                    "SearchNext",
                    "SearchNext",
                    typeof(CustomCommand),
                    null
            );
        public static readonly RoutedUICommand SearchPrevious = new RoutedUICommand
            (
                    "SearchPrevious",
                    "SearchPrevious",
                    typeof(CustomCommand),
                    null
            );
    }
}

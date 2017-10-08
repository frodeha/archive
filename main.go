package main

import (
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"os/user"
	"path"
	"path/filepath"
	"strings"

	"github.com/frodeha/archive/archive"
	"github.com/spf13/cobra"
)

var (
	folder string
)

var composer string
var arranger string
var force bool

func sync(cmd *cobra.Command, args []string) error {

	key, err := openAndLoadCurrentFile()
	if err != nil {
		return err
	}

	filepath.Walk(folder, func(p string, info os.FileInfo, err error) error {
		if p == folder {
			return nil
		}

		if info.IsDir() {
			name := info.Name()
			key.AddRow(name, "", "",
				exists(p, name, "partitur"),
				exists(p, name, "treblås"),
				exists(p, name, "messing"),
				exists(p, name, "slagverk"))
		}

		return nil
	})

	key2 := Archive.NewArchiveKey("")

	for _, v := range key.Rows() {
		key2.AddRow(v.Name, v.Composer, v.Arranger, v.HasScore, v.HasWoodwind, v.HasBrass, v.HasPercussion)
	}

	return backupAndOverwriteCurrentFile(key2)
}

func switchArchive(cmd *cobra.Command, args []string) error {
	key := Archive.NewArchiveKey("")
	err := key.SaveAs(currentDirectoryFile())
	if err != nil {
		return err
	}

	config := fmt.Sprintf("location = %s", args[0])
	return ioutil.WriteFile(".archive", []byte(config), 0644)
}

func add(cmd *cobra.Command, args []string) error {

	location := args[0]
	name := path.Base(location)

	key, err := openAndLoadCurrentFile()
	if err != nil {
		return err
	}

	key.AddRow(name,
		composer,
		arranger,
		exists(location, name, "partitur"),
		exists(location, name, "treblås"),
		exists(location, name, "messing"),
		exists(location, name, "slagverk"),
	)

	err = key.Save()
	if err != nil {
		return err
	}

	destinationDirectory := fmt.Sprintf("%s/%s", folder, name)
	err = os.Mkdir(destinationDirectory, 0755)
	if err != nil {
		return err
	}

	copyFile(fmt.Sprintf("%s/%s - partitur.pdf", location, name), fmt.Sprintf("%s/%s - partitur.pdf", destinationDirectory, name))
	copyFile(fmt.Sprintf("%s/%s - treblås.pdf", location, name), fmt.Sprintf("%s/%s - treblås.pdf", destinationDirectory, name))
	copyFile(fmt.Sprintf("%s/%s - messing.pdf", location, name), fmt.Sprintf("%s/%s - messing.pdf", destinationDirectory, name))
	copyFile(fmt.Sprintf("%s/%s - slagverk.pdf", location, name), fmt.Sprintf("%s/%s - slagverk.pdf", destinationDirectory, name))

	return nil
}

func update(cmd *cobra.Command, args []string) error {
	name := args[0]
	key, err := openAndLoadCurrentFile()
	if err != nil {
		return err
	}

	row, err := key.GetRow(name)
	if err != nil {
		return err
	}

	location := fmt.Sprintf("%s/%s", folder, name)
	if force {
		row.Composer = composer
		row.Arranger = arranger
	} else {
		if composer != "" {
			row.Composer = composer
		}

		if arranger != "" {
			row.Arranger = arranger
		}
	}

	row.HasScore = exists(location, name, "partitur")
	row.HasWoodwind = exists(location, name, "treblås")
	row.HasBrass = exists(location, name, "messing")
	row.HasPercussion = exists(location, name, "slagverk")

	err = key.UpdateRow(row)
	if err != nil {
		return err
	}

	return key.Save()
}

func delete(cmd *cobra.Command, args []string) error {
	name := args[0]
	key, err := openAndLoadCurrentFile()
	if err != nil {
		return err
	}

	err = key.DeleteRow(name)
	if err != nil {
		return err
	}

	return key.Save()
}

func print(cmd *cobra.Command, args []string) error {

	key, err := openAndLoadCurrentFile()
	if err != nil {
		return err
	}

	name := ""
	if len(args) > 0 {
		name = args[0]
	}

	key.Print(name)

	return nil
}

func main() {

	folder = loadConfig()

	RootCmd := &cobra.Command{
		Use:   "archive",
		Short: "Archive is an archive",
		Long:  fmt.Sprintf(`Archive allows you to structure files across folders. Archive is currently configured to use the folder: %s`, folder),
	}

	RootCmd.AddCommand(&cobra.Command{
		Use:   "version",
		Short: "Print the archive version",
		Long:  ``,
		Run: func(cmd *cobra.Command, args []string) {
			fmt.Printf("Version: 1.0.0\n")
		},
	})

	RootCmd.AddCommand(&cobra.Command{
		Use:   "use [path]",
		Short: "Set the active archive",
		Long:  `Set the active archive folder. The folder must be an absolute path`,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return fmt.Errorf("Missing path to directory")
			}

			info, err := os.Stat(args[0])
			if err != nil {
				return fmt.Errorf("Error while checking folder, %v", err)
			}

			if !info.IsDir() {
				return fmt.Errorf("%s is not a directory", args[0])
			}

			return nil
		},
		RunE: switchArchive,
	})

	RootCmd.AddCommand(&cobra.Command{
		Use:   "print",
		Short: "Print the archive overview",
		Long:  `Set the active archive folder. The folder must be an absolute path`,
		RunE:  print,
	})

	addCommand := &cobra.Command{
		Use:   "add",
		Short: "Adds new piece",
		Long:  ``,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return fmt.Errorf("Missing path to folder")
			}

			info, err := os.Stat(args[0])
			if err != nil {
				return fmt.Errorf("Error while checking folder, %v", err)
			}

			if !info.IsDir() {
				return fmt.Errorf("%s is not a directory", args[0])
			}

			return nil
		},
		RunE: add,
	}

	addCommand.Flags().StringVarP(&composer, "composer", "c", "", "Set the composer of the piece")
	addCommand.Flags().StringVarP(&arranger, "arranger", "a", "", "Set the arranger of the piece")

	RootCmd.AddCommand(addCommand)

	updateCommand := &cobra.Command{
		Use:   "update",
		Short: "Updates an existing piece",
		Long:  ``,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return fmt.Errorf("Missing name")
			}

			return nil
		},
		RunE: update,
	}

	updateCommand.Flags().StringVarP(&composer, "composer", "c", "", "Set the composer of the piece")
	updateCommand.Flags().StringVarP(&arranger, "arranger", "a", "", "Set the arranger of the piece")
	updateCommand.Flags().BoolVarP(&force, "force", "f", false, "Force update all properties")

	RootCmd.AddCommand(updateCommand)

	RootCmd.AddCommand(&cobra.Command{
		Use:   "delete",
		Short: "syncs the current folder and the excel file",
		Long:  ``,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return fmt.Errorf("Missing name")
			}

			return nil
		},
		RunE: delete,
	})

	RootCmd.AddCommand(&cobra.Command{
		Use:   "sync",
		Short: "syncs the current folder and the excel file",
		Long:  ``,
		RunE:  sync,
	})

	if err := RootCmd.Execute(); err != nil {
		fmt.Printf("%v\n", err)
		os.Exit(1)
	}
}

func loadConfig() string {
	usr, err := user.Current()
	if err != nil {
		log.Fatal(err)
	}

	data, err := ioutil.ReadFile(fmt.Sprintf("%s/.archive", usr.HomeDir))

	if err != nil {
		return "/tmp"
	}

	conf := string(data)
	parts := strings.Split(conf, " = ")

	if len(parts) != 2 {
		return "/tmp"
	}

	return parts[1]
}

func exists(path, name, part string) bool {
	_, err := os.Stat(fmt.Sprintf("%s/%s - %s.pdf", path, name, part))
	if err != nil {
		return false
	}
	return true
}

func currentDirectoryFile() string {
	return fmt.Sprintf("%s/_archive.xlsx", folder)
}

func copyFile(source, destination string) {

	src, err := os.Open(source)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer src.Close()

	dest, err := os.Create(destination)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer dest.Close()

	_, err = io.Copy(dest, src)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Printf("Copied %s\n", source)
}

func backupAndOverwriteCurrentFile(a *Archive.ArchiveKey) error {
	copyFile(currentDirectoryFile(), fmt.Sprintf("%s.bak", currentDirectoryFile()))
	return a.SaveAs(currentDirectoryFile())
}

func openAndLoadCurrentFile() (*Archive.ArchiveKey, error) {
	key := Archive.NewArchiveKey(currentDirectoryFile())
	err := key.Load()
	if err != nil {
		return nil, err
	}

	return key, nil
}
